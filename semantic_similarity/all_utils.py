"""Minimal whitening utilities used by ``similarity_computation.py``.

The original project imported these functions from an external source tree.
Keeping the two required functions here makes the analysis self-contained.
"""

from __future__ import annotations

import pickle
from pathlib import Path
from typing import Any

import numpy as np
import torch


def load_whiten(path: str | Path) -> tuple[np.ndarray, np.ndarray]:
    """Load a whitening kernel and bias from a pickle file."""
    whitening_path = Path(path).expanduser().resolve()
    if not whitening_path.is_file():
        raise FileNotFoundError(f"Whitening file not found: {whitening_path}")

    with whitening_path.open("rb") as stream:
        parameters: dict[str, Any] = pickle.load(stream)

    if "kernel" not in parameters or "bias" not in parameters:
        raise ValueError(
            f"Whitening file must contain 'kernel' and 'bias': {whitening_path}"
        )

    kernel = np.asarray(parameters["kernel"])
    bias = np.asarray(parameters["bias"])
    if kernel.ndim != 2 or bias.shape[-1] != kernel.shape[0]:
        raise ValueError(
            f"Incompatible whitening shapes: kernel={kernel.shape}, bias={bias.shape}"
        )
    return kernel, bias


def transform_and_normalize(
    vectors: torch.Tensor,
    kernel: np.ndarray,
    bias: np.ndarray,
    *,
    normalization_dim: int = 1,
) -> torch.Tensor:
    """Apply the whitening transform and L2-normalize along one dimension."""
    device = vectors.device
    dtype = vectors.dtype
    kernel_tensor = torch.as_tensor(kernel, dtype=dtype, device=device)
    bias_tensor = torch.as_tensor(bias, dtype=dtype, device=device)

    transformed = torch.matmul(vectors + bias_tensor, kernel_tensor)
    denominator = torch.linalg.vector_norm(
        transformed, ord=2, dim=normalization_dim, keepdim=True
    ).clamp_min(torch.finfo(dtype).eps)
    return transformed / denominator
