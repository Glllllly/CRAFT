from __future__ import annotations

import importlib.util
import json
import socket
import socketserver
import threading
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable


ROOT = Path(__file__).resolve().parents[1]
VISION_DIR = ROOT / "vision"
INFER_LORA_PATH = VISION_DIR / "lora_infer.py"


def _import_module_from_path(module_name: str, path: Path):
    spec = importlib.util.spec_from_file_location(module_name, str(path))
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to import module from {path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def load_lora_once(model_name: str, lora_path: str, device_map: str):
    import torch
    from peft import PeftModel
    from transformers import AutoModelForVision2Seq, AutoProcessor

    try:
        from transformers import Qwen3VLForConditionalGeneration
    except Exception:
        Qwen3VLForConditionalGeneration = None

    infer_lora_mod = _import_module_from_path("infer_lora_reuse_mod", INFER_LORA_PATH)
    processor = AutoProcessor.from_pretrained(model_name, trust_remote_code=True)
    dtype = torch.bfloat16 if torch.cuda.is_available() else torch.float32
    if Qwen3VLForConditionalGeneration is not None:
        base = Qwen3VLForConditionalGeneration.from_pretrained(
            model_name,
            dtype=dtype,
            trust_remote_code=True,
            device_map=device_map,
        )
    else:
        base = AutoModelForVision2Seq.from_pretrained(
            model_name,
            dtype=dtype,
            trust_remote_code=True,
            device_map=device_map,
        )
    model = PeftModel.from_pretrained(base, lora_path)
    model.eval()
    model_device = next(model.parameters()).device
    return infer_lora_mod, processor, model, model_device


def infer_lora_inprocess(
    infer_lora_mod,
    processor,
    model,
    model_device,
    image_path: Path,
    slots_json: Path,
    out_json: Path,
    max_new_tokens: int,
    temperature: float,
    slots_chunk_size: int,
    crop_by_slot_chunk: bool,
    crop_left_pad: int,
    crop_top_pad: int,
    crop_right_pad: int,
    crop_bottom_pad: int,
    max_image_size: int,
    resize_after_crop: bool,
    annotate_slot_id: bool,
    prompt_context_text: str = "",
    debug_dir: str = "",
) -> None:
    infer_lora_mod.run_inference(
        processor=processor,
        model=model,
        model_device=model_device,
        image_path=str(image_path),
        slots_json_path=str(slots_json),
        output_json_path=str(out_json),
        max_new_tokens=int(max_new_tokens),
        temperature=float(temperature),
        slots_chunk_size=int(slots_chunk_size),
        crop_by_slot_chunk=bool(crop_by_slot_chunk),
        crop_left_pad=int(crop_left_pad),
        crop_top_pad=int(crop_top_pad),
        crop_right_pad=int(crop_right_pad),
        crop_bottom_pad=int(crop_bottom_pad),
        max_image_size=int(max_image_size),
        resize_after_crop=bool(resize_after_crop),
        annotate_slot_id=bool(annotate_slot_id),
        prompt_context_text=str(prompt_context_text or ""),
        debug_dir=str(debug_dir or ""),
    )


@dataclass
class _ServerState:
    infer_lora_mod: Any
    processor: Any
    model: Any
    model_device: Any
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
            elif cmd == "infer":
                state: _ServerState = self.server.state
                infer_lora_inprocess(
                    infer_lora_mod=state.infer_lora_mod,
                    processor=state.processor,
                    model=state.model,
                    model_device=state.model_device,
                    image_path=Path(str(req["image_path"])),
                    slots_json=Path(str(req["slots_json"])),
                    out_json=Path(str(req["output_json"])),
                    max_new_tokens=int(req["max_new_tokens"]),
                    temperature=float(req["temperature"]),
                    slots_chunk_size=int(req["slots_chunk_size"]),
                    crop_by_slot_chunk=bool(req["crop_by_slot_chunk"]),
                    crop_left_pad=int(req["crop_left_pad"]),
                    crop_top_pad=int(req["crop_top_pad"]),
                    crop_right_pad=int(req["crop_right_pad"]),
                    crop_bottom_pad=int(req["crop_bottom_pad"]),
                    max_image_size=int(req["max_image_size"]),
                    resize_after_crop=bool(req["resize_after_crop"]),
                    annotate_slot_id=bool(req["annotate_slot_id"]),
                    prompt_context_text=str(req.get("prompt_context_text") or ""),
                    debug_dir=str(req.get("debug_dir") or ""),
                )
                resp = {"ok": True, "output_json": str(req["output_json"])}
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


class QwenLoraReuseServer:
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


def start_lora_server(
    model_name: str,
    lora_path: str,
    device_map: str,
    host: str = "127.0.0.1",
    port: int = 0,
    log_fn: Callable[[str], None] | None = None,
) -> QwenLoraReuseServer:
    log = log_fn or (lambda *_args, **_kwargs: None)
    log(f"[QWEN_REUSE] loading model={model_name} lora={lora_path}")
    infer_lora_mod, processor, model, model_device = load_lora_once(model_name, lora_path, device_map)
    server = _ThreadingTCPServer((host, int(port)), _RequestHandler)
    server.state = _ServerState(
        infer_lora_mod=infer_lora_mod,
        processor=processor,
        model=model,
        model_device=model_device,
        log_fn=log,
    )
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    log(f"[QWEN_REUSE] ready addr={server.server_address[0]}:{server.server_address[1]}")
    return QwenLoraReuseServer(server, thread)


start_qwen_lora_reuse_server = start_lora_server


def _parse_address(address: str) -> tuple[str, int]:
    host, sep, port_s = str(address or "").strip().rpartition(":")
    if not sep or not host or not port_s:
        raise ValueError(f"invalid Qwen reuse address: {address!r}")
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
        raise RuntimeError("empty response from Qwen reuse server")
    return json.loads(buff.decode("utf-8").strip())


def request_infer_lora_from_server(
    address: str,
    image_path: Path,
    slots_json: Path,
    output_json: Path,
    max_new_tokens: int,
    temperature: float,
    slots_chunk_size: int,
    crop_by_slot_chunk: bool,
    crop_left_pad: int,
    crop_top_pad: int,
    crop_right_pad: int,
    crop_bottom_pad: int,
    max_image_size: int,
    resize_after_crop: bool,
    annotate_slot_id: bool,
    prompt_context_text: str = "",
    debug_dir: Path | None = None,
    timeout_s: float = 3600.0,
) -> Path:
    payload = {
        "cmd": "infer",
        "image_path": str(image_path),
        "slots_json": str(slots_json),
        "output_json": str(output_json),
        "max_new_tokens": int(max_new_tokens),
        "temperature": float(temperature),
        "slots_chunk_size": int(slots_chunk_size),
        "crop_by_slot_chunk": bool(crop_by_slot_chunk),
        "crop_left_pad": int(crop_left_pad),
        "crop_top_pad": int(crop_top_pad),
        "crop_right_pad": int(crop_right_pad),
        "crop_bottom_pad": int(crop_bottom_pad),
        "max_image_size": int(max_image_size),
        "resize_after_crop": bool(resize_after_crop),
        "annotate_slot_id": bool(annotate_slot_id),
        "prompt_context_text": str(prompt_context_text or ""),
        "debug_dir": str(debug_dir) if debug_dir is not None else "",
    }
    resp = _request(address=address, payload=payload, timeout_s=timeout_s)
    if not bool(resp.get("ok")):
        raise RuntimeError(str(resp.get("error") or "Qwen reuse infer failed"))
    return output_json
