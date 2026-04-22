import math
from collections import deque
from pathlib import Path
from typing import Dict, List, Tuple

import numpy as np
import torch
import torch.nn as nn
import torch.nn.functional as F
from transformers import SegformerConfig, SegformerForSemanticSegmentation

DEFAULT_PRETRAINED_MODEL = "nvidia/segformer-b1-finetuned-ade-512-512"
DEFAULT_FINAL_LOSS_WEIGHT = 1.0
DEFAULT_HORIZONTAL_LOSS_WEIGHT = 0.3
DEFAULT_VERTICAL_LOSS_WEIGHT = 0.3


def connected_components(mask: np.ndarray) -> List[List[Tuple[int, int]]]:
    height, width = mask.shape
    visited = np.zeros_like(mask, dtype=bool)
    components: List[List[Tuple[int, int]]] = []

    for y in range(height):
        for x in range(width):
            if mask[y, x] == 0 or visited[y, x]:
                continue

            queue = deque([(y, x)])
            visited[y, x] = True
            pixels: List[Tuple[int, int]] = []

            while queue:
                cy, cx = queue.popleft()
                pixels.append((cy, cx))

                for ny, nx in ((cy - 1, cx), (cy + 1, cx), (cy, cx - 1), (cy, cx + 1)):
                    if ny < 0 or ny >= height or nx < 0 or nx >= width:
                        continue
                    if visited[ny, nx] or mask[ny, nx] == 0:
                        continue
                    visited[ny, nx] = True
                    queue.append((ny, nx))

            components.append(pixels)

    return components


def build_rectangle_targets(mask: np.ndarray) -> Tuple[np.ndarray, np.ndarray, np.ndarray]:
    binary_mask = (mask > 0).astype(np.uint8)
    rect_mask = np.zeros_like(binary_mask, dtype=np.float32)
    horizontal_edges = np.zeros_like(binary_mask, dtype=np.float32)
    vertical_edges = np.zeros_like(binary_mask, dtype=np.float32)

    for pixels in connected_components(binary_mask):
        ys = [y for y, _ in pixels]
        xs = [x for _, x in pixels]
        y0, y1 = min(ys), max(ys)
        x0, x1 = min(xs), max(xs)

        rect_mask[y0 : y1 + 1, x0 : x1 + 1] = 1.0
        horizontal_edges[y0, x0 : x1 + 1] = 1.0
        horizontal_edges[y1, x0 : x1 + 1] = 1.0
        vertical_edges[y0 : y1 + 1, x0] = 1.0
        vertical_edges[y0 : y1 + 1, x1] = 1.0

    return rect_mask, horizontal_edges, vertical_edges


class RectangleAwareSegformer(nn.Module):
    def __init__(
        self,
        pretrained_name: str = DEFAULT_PRETRAINED_MODEL,
        init_mode: str = "pretrained",
    ):
        super().__init__()
        if init_mode == "pretrained":
            base_model = SegformerForSemanticSegmentation.from_pretrained(
                pretrained_name,
                num_labels=1,
                ignore_mismatched_sizes=True,
            )
        elif init_mode == "scratch":
            config = SegformerConfig.from_pretrained(
                pretrained_name,
                num_labels=1,
            )
            base_model = SegformerForSemanticSegmentation(config)
        else:
            raise ValueError(f"Unsupported init_mode: {init_mode}")

        self.segformer = base_model.segformer
        self.decode_head = base_model.decode_head

        decoder_channels = self.decode_head.config.decoder_hidden_size
        self.horizontal_head = nn.Sequential(
            nn.Conv2d(decoder_channels, decoder_channels, kernel_size=(1, 7), padding=(0, 3)),
            nn.BatchNorm2d(decoder_channels),
            nn.ReLU(),
            nn.Conv2d(decoder_channels, 1, kernel_size=1),
        )
        self.vertical_head = nn.Sequential(
            nn.Conv2d(decoder_channels, decoder_channels, kernel_size=(7, 1), padding=(3, 0)),
            nn.BatchNorm2d(decoder_channels),
            nn.ReLU(),
            nn.Conv2d(decoder_channels, 1, kernel_size=1),
        )
        self.fusion_head = nn.Sequential(
            nn.Conv2d(decoder_channels + 2, decoder_channels, kernel_size=1),
            nn.BatchNorm2d(decoder_channels),
            nn.ReLU(),
            nn.Conv2d(decoder_channels, 1, kernel_size=1),
        )

    def _reshape_encoder_hidden_state(
        self,
        encoder_hidden_state: torch.Tensor,
        batch_size: int,
    ) -> torch.Tensor:
        if self.decode_head.config.reshape_last_stage is False and encoder_hidden_state.ndim == 3:
            height = width = int(math.sqrt(encoder_hidden_state.shape[-1]))
            encoder_hidden_state = (
                encoder_hidden_state.reshape(batch_size, height, width, -1).permute(0, 3, 1, 2).contiguous()
            )
        return encoder_hidden_state

    def _decode_features(self, encoder_hidden_states: Tuple[torch.Tensor, ...]) -> torch.Tensor:
        batch_size = encoder_hidden_states[-1].shape[0]
        all_hidden_states = []

        for encoder_hidden_state, mlp in zip(encoder_hidden_states, self.decode_head.linear_c):
            encoder_hidden_state = self._reshape_encoder_hidden_state(encoder_hidden_state, batch_size)
            height, width = encoder_hidden_state.shape[2], encoder_hidden_state.shape[3]
            hidden_state = mlp(encoder_hidden_state)
            hidden_state = hidden_state.permute(0, 2, 1).reshape(batch_size, -1, height, width)
            hidden_state = F.interpolate(
                hidden_state,
                size=encoder_hidden_states[0].size()[2:],
                mode="bilinear",
                align_corners=False,
            )
            all_hidden_states.append(hidden_state)

        fused = self.decode_head.linear_fuse(torch.cat(all_hidden_states[::-1], dim=1))
        fused = self.decode_head.batch_norm(fused)
        fused = self.decode_head.activation(fused)
        fused = self.decode_head.dropout(fused)
        return fused

    def forward(self, pixel_values: torch.Tensor) -> Dict[str, torch.Tensor]:
        outputs = self.segformer(
            pixel_values,
            output_hidden_states=True,
            return_dict=True,
        )
        decoder_features = self._decode_features(outputs.hidden_states)
        horizontal_logits = self.horizontal_head(decoder_features)
        vertical_logits = self.vertical_head(decoder_features)
        final_logits = self.fusion_head(
            torch.cat([decoder_features, horizontal_logits, vertical_logits], dim=1)
        )
        return {
            "final_logits": final_logits,
            "horizontal_logits": horizontal_logits,
            "vertical_logits": vertical_logits,
        }


def compute_rectangle_losses(
    outputs: Dict[str, torch.Tensor],
    rect_target: torch.Tensor,
    horizontal_target: torch.Tensor,
    vertical_target: torch.Tensor,
    final_loss_weight: float = DEFAULT_FINAL_LOSS_WEIGHT,
    horizontal_loss_weight: float = DEFAULT_HORIZONTAL_LOSS_WEIGHT,
    vertical_loss_weight: float = DEFAULT_VERTICAL_LOSS_WEIGHT,
) -> Dict[str, torch.Tensor]:
    if rect_target.ndim == 3:
        rect_target = rect_target.unsqueeze(1)
    if horizontal_target.ndim == 3:
        horizontal_target = horizontal_target.unsqueeze(1)
    if vertical_target.ndim == 3:
        vertical_target = vertical_target.unsqueeze(1)

    final_logits = F.interpolate(
        outputs["final_logits"],
        size=rect_target.shape[-2:],
        mode="bilinear",
        align_corners=False,
    )
    horizontal_logits = F.interpolate(
        outputs["horizontal_logits"],
        size=horizontal_target.shape[-2:],
        mode="bilinear",
        align_corners=False,
    )
    vertical_logits = F.interpolate(
        outputs["vertical_logits"],
        size=vertical_target.shape[-2:],
        mode="bilinear",
        align_corners=False,
    )

    final_loss = F.binary_cross_entropy_with_logits(final_logits, rect_target)
    horizontal_loss = F.binary_cross_entropy_with_logits(horizontal_logits, horizontal_target)
    vertical_loss = F.binary_cross_entropy_with_logits(vertical_logits, vertical_target)
    total_loss = (
        final_loss_weight * final_loss
        + horizontal_loss_weight * horizontal_loss
        + vertical_loss_weight * vertical_loss
    )

    return {
        "total": total_loss,
        "final": final_loss,
        "horizontal": horizontal_loss,
        "vertical": vertical_loss,
        "final_logits": final_logits,
        "horizontal_logits": horizontal_logits,
        "vertical_logits": vertical_logits,
    }


def load_flexible_checkpoint(
    model: torch.nn.Module,
    weight_path: Path,
    map_location: torch.device | str | None = None,
) -> Dict[str, int]:
    payload = torch.load(weight_path, map_location=map_location)
    if isinstance(payload, dict) and "state_dict" in payload and isinstance(payload["state_dict"], dict):
        state = payload["state_dict"]
    else:
        state = payload

    normalized = {}
    for key, value in state.items():
        new_key = key
        if new_key.startswith("module."):
            new_key = new_key[len("module.") :]
        if new_key.startswith("seg."):
            new_key = new_key[len("seg.") :]
        normalized[new_key] = value

    model_state = model.state_dict()
    filtered = {}
    skipped = 0
    for key, value in normalized.items():
        if key in model_state and model_state[key].shape == value.shape:
            filtered[key] = value
        else:
            skipped += 1

    missing, unexpected = model.load_state_dict(filtered, strict=False)
    return {
        "loaded": len(filtered),
        "skipped": skipped,
        "missing": len(missing),
        "unexpected": len(unexpected),
    }
