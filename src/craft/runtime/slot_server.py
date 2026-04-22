from __future__ import annotations

import importlib.util
import json
import socket
import socketserver
import sys
import threading
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

import openpyxl
from openai import OpenAI
from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
VISION_DIR = ROOT / "vision"
SLOT_DETECT_PATH = VISION_DIR / "slot_detect.py"


def _import_module_from_path(module_name: str, path: Path):
    vision_dir = str(VISION_DIR.resolve())
    if vision_dir not in sys.path:
        sys.path.insert(0, vision_dir)
    spec = importlib.util.spec_from_file_location(module_name, str(path))
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to import module from {path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def load_slot_detector_once(ckpt_path: str):
    qwen_small_mod = _import_module_from_path("slot_server_mod", SLOT_DETECT_PATH)
    device = "cuda" if qwen_small_mod.torch.cuda.is_available() else "cpu"
    seg = qwen_small_mod.RectangleAwareSegformer(
        pretrained_name=qwen_small_mod.DEFAULT_PRETRAINED_MODEL,
        init_mode="pretrained",
    ).to(device)
    load_stats = qwen_small_mod.load_flexible_checkpoint(seg, ckpt_path, map_location="cpu")
    seg.eval()
    return qwen_small_mod, seg, device, load_stats


def detect_slots_inprocess(
    qwen_small_mod,
    seg_model,
    seg_device: str,
    image_path: Path,
    xlsx_path: Path,
    out_dir: Path,
    sheet_name: str | None,
    edges_json: Path | None,
    bounds_json: Path | None,
    qwen_base_url: str,
    qwen_model: str,
    qwen_api_key: str,
    skip_qwen: bool,
) -> dict[str, str]:
    out_dir.mkdir(parents=True, exist_ok=True)

    mask01_hw = qwen_small_mod.seg_infer_mask(seg_model, str(image_path), device=seg_device, img_size=512, prob_thresh=-1.0)

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    try:
        ws = wb[sheet_name] if sheet_name else wb.active

        if bounds_json is not None and bounds_json.exists():
            b = json.loads(bounds_json.read_text(encoding="utf-8"))
            min_row = int(b["min_row"])
            min_col = int(b["min_col"])
            max_row = int(b["max_row"])
            max_col = int(b["max_col"])
        else:
            min_row, min_col, max_row, max_col = qwen_small_mod.get_used_bbox(ws)

        if edges_json is not None and edges_json.exists():
            e = json.loads(edges_json.read_text(encoding="utf-8"))
            x_edges = e.get("x_edges")
            y_edges = e.get("y_edges")
        else:
            x_edges, y_edges = qwen_small_mod.compute_edges(ws, min_row, min_col, max_row, max_col, dpi=96)

        img = Image.open(image_path)
        W, H = img.size
        x_edges, y_edges = qwen_small_mod.cluster_edges_to_image(x_edges, y_edges, W, H)
        slots = qwen_small_mod.mask_to_excel_slots(
            mask01_hw=mask01_hw,
            ws=ws,
            x_edges=x_edges,
            y_edges=y_edges,
            min_row=min_row,
            min_col=min_col,
            max_row=max_row,
            max_col=max_col,
            min_area=200,
        )
    finally:
        wb.close()

    mask_path = out_dir / "mask.png"
    overlay_path = out_dir / "overlay.png"
    slots_path = out_dir / "slots_aligned.json"
    qwen_pairs_path = out_dir / "qwen_pairs.json"

    qwen_small_mod.cv2.imwrite(str(mask_path), (mask01_hw * 255).astype(qwen_small_mod.np.uint8))
    qwen_small_mod.draw_overlay(
        str(image_path),
        mask01_hw,
        slots,
        str(overlay_path),
        label_mode="id",
        min_label_box_area=300,
    )
    slots_payload = {
        "image": str(image_path),
        "xlsx": str(xlsx_path),
        "sheet": sheet_name or ws.title,
        "image_size": {"w": W, "h": H},
        "used_range": {"min_row": min_row, "min_col": min_col, "max_row": max_row, "max_col": max_col},
        "slots": slots,
    }
    slots_path.write_text(json.dumps(slots_payload, ensure_ascii=False, indent=2), encoding="utf-8")

    if not skip_qwen:
        client = OpenAI(api_key=qwen_api_key, base_url=qwen_base_url, timeout=3600)
        qwen_out = qwen_small_mod.run_qwen_batches(
            client=client,
            model_name=qwen_model,
            overlay_img_path=str(overlay_path),
            slots=slots,
            batch_size=10,
            temperature=0.0,
            out_dir=str(out_dir),
        )
        qwen_pairs_path.write_text(json.dumps(qwen_out, ensure_ascii=False, indent=2), encoding="utf-8")

    return {
        "mask_png": str(mask_path),
        "overlay_png": str(overlay_path),
        "slots_json": str(slots_path),
        "qwen_pairs": str(qwen_pairs_path),
    }


@dataclass
class _ServerState:
    qwen_small_mod: Any
    seg_model: Any
    seg_device: str
    log_fn: Callable[[str], None]


class _ThreadingTCPServer(socketserver.ThreadingMixIn, socketserver.TCPServer):
    allow_reuse_address = True
    daemon_threads = True


class _RequestHandler(socketserver.StreamRequestHandler):
    def handle(self) -> None:
        raw = self.rfile.readline()
        if not raw:
            return
        try:
            req = json.loads(raw.decode("utf-8"))
            cmd = str(req.get("cmd") or "").strip().lower()
            if cmd == "ping":
                resp = {"ok": True, "pong": True}
            elif cmd == "detect":
                state: _ServerState = self.server.state
                outputs = detect_slots_inprocess(
                    qwen_small_mod=state.qwen_small_mod,
                    seg_model=state.seg_model,
                    seg_device=state.seg_device,
                    image_path=Path(str(req["image_path"])),
                    xlsx_path=Path(str(req["xlsx_path"])),
                    out_dir=Path(str(req["out_dir"])),
                    sheet_name=str(req.get("sheet_name") or "").strip() or None,
                    edges_json=(Path(str(req["edges_json"])) if str(req.get("edges_json") or "").strip() else None),
                    bounds_json=(Path(str(req["bounds_json"])) if str(req.get("bounds_json") or "").strip() else None),
                    qwen_base_url=str(req.get("qwen_base_url") or ""),
                    qwen_model=str(req.get("qwen_model") or ""),
                    qwen_api_key=str(req.get("qwen_api_key") or ""),
                    skip_qwen=bool(req.get("skip_qwen")),
                )
                resp = {"ok": True, "outputs": outputs}
            else:
                resp = {"ok": False, "error": f"unsupported cmd: {cmd}"}
        except Exception as exc:
            resp = {
                "ok": False,
                "error": str(exc),
                "traceback": traceback.format_exc(),
            }
        self.wfile.write((json.dumps(resp, ensure_ascii=False) + "\n").encode("utf-8"))
        self.wfile.flush()


class SlotDetectReuseServer:
    def __init__(self, server: _ThreadingTCPServer, thread: threading.Thread):
        self._server = server
        self._thread = thread

    @property
    def host(self) -> str:
        return str(self._server.server_address[0])

    @property
    def port(self) -> int:
        return int(self._server.server_address[1])

    @property
    def address(self) -> str:
        return f"{self.host}:{self.port}"

    def close(self) -> None:
        self._server.shutdown()
        self._server.server_close()
        self._thread.join(timeout=5.0)


def start_slot_server(
    ckpt_path: str,
    host: str = "127.0.0.1",
    port: int = 0,
    log_fn: Callable[[str], None] | None = None,
) -> SlotDetectReuseServer:
    log = log_fn or (lambda *_args, **_kwargs: None)
    log(f"[SLOT_REUSE] loading ckpt={ckpt_path}")
    qwen_small_mod, seg_model, seg_device, load_stats = load_slot_detector_once(ckpt_path)
    log(
        f"[SLOT_REUSE] loaded loaded={load_stats['loaded']} skipped={load_stats['skipped']} "
        f"missing={load_stats['missing']} unexpected={load_stats['unexpected']}"
    )
    server = _ThreadingTCPServer((host, int(port)), _RequestHandler)
    server.state = _ServerState(
        qwen_small_mod=qwen_small_mod,
        seg_model=seg_model,
        seg_device=seg_device,
        log_fn=log,
    )
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    log(f"[SLOT_REUSE] ready addr={server.server_address[0]}:{server.server_address[1]}")
    return SlotDetectReuseServer(server, thread)


start_slot_detect_reuse_server = start_slot_server


def _parse_address(address: str) -> tuple[str, int]:
    host, sep, port_s = str(address or "").strip().rpartition(":")
    if not sep or not host or not port_s:
        raise ValueError(f"invalid slot-detect reuse address: {address!r}")
    return host, int(port_s)


def _request(address: str, payload: dict[str, Any], timeout_s: float = 3600.0) -> dict[str, Any]:
    host, port = _parse_address(address)
    with socket.create_connection((host, port), timeout=timeout_s) as sock:
        sock.sendall((json.dumps(payload, ensure_ascii=False) + "\n").encode("utf-8"))
        buff = b""
        while not buff.endswith(b"\n"):
            chunk = sock.recv(65536)
            if not chunk:
                break
            buff += chunk
    if not buff:
        raise RuntimeError("empty response from slot-detect reuse server")
    return json.loads(buff.decode("utf-8").strip())


def request_slot_detect_from_server(
    address: str,
    image_path: Path,
    xlsx_path: Path,
    out_dir: Path,
    sheet_name: str | None,
    edges_json: Path | None,
    bounds_json: Path | None,
    qwen_base_url: str,
    qwen_model: str,
    qwen_api_key: str,
    skip_qwen: bool,
    timeout_s: float = 3600.0,
) -> dict[str, Path]:
    payload = {
        "cmd": "detect",
        "image_path": str(image_path),
        "xlsx_path": str(xlsx_path),
        "out_dir": str(out_dir),
        "sheet_name": str(sheet_name or ""),
        "edges_json": str(edges_json) if edges_json is not None else "",
        "bounds_json": str(bounds_json) if bounds_json is not None else "",
        "qwen_base_url": str(qwen_base_url or ""),
        "qwen_model": str(qwen_model or ""),
        "qwen_api_key": str(qwen_api_key or ""),
        "skip_qwen": bool(skip_qwen),
    }
    resp = _request(address=address, payload=payload, timeout_s=timeout_s)
    if not bool(resp.get("ok")):
        raise RuntimeError(str(resp.get("error") or "slot-detect reuse failed"))
    outputs = resp.get("outputs") or {}
    return {str(k): Path(str(v)) for k, v in outputs.items()}
import sys
