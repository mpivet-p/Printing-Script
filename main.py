"""
Saturn Card Printer PRN decoder / encoder.

PRN file structure (reverse-engineered):
    ESC SOJ=<n>,<job_name> CR
    ESC SOP=<n>,1 CR
    ESC PDM=2512,<2512 bytes TFSM header>
    CR ESC SCP=80,<80 bytes color-panel config>
    CR ESC SCB=1954720,<color pixel data>
    CR ESC SKP=80,<80 bytes K-panel config>
    CR ESC SKB=1954720,<K-resin pixel data>
    CR ESC EOP=<n>,1 CR
    ESC EOJ=<n>,1 CR

Color pixel data: 8-bit-per-channel YMC, interleaved byte-by-byte:
    Y0 M0 C0  Y1 M1 C1  ...  Y1012 M1012 C1012  PAD
Row stride = 3040 bytes  (1013 * 3 + 1 padding 0xFF).
643 rows for a standard CR-80 card at 300 DPI.

K-resin pixel data: same layout (3-byte interleaved, 3040 stride),
all 3 channels carry the same value: 0 = K-resin, 255 = no K.

Subtractive ink model: 255 = no ink (white), 0 = full ink.
Where K is active, the color layer is set to 255 (no dye).
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import numpy as np
from PIL import Image

CARD_WIDTH = 1013
CARD_HEIGHT = 643
ROW_STRIDE = 3040  # 1013 * 3 + 1 padding byte
PIXEL_DATA_SIZE = ROW_STRIDE * CARD_HEIGHT  # 1_954_720

SCB_MARKER = b"\x1bSCB="
SKB_MARKER = b"\x1bSKB="
DEFAULT_TEMPLATE = Path(__file__).parent / "another-test.prn"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _find_data_after_marker(data: bytes, marker: bytes) -> int:
    """Return the byte offset just after '<marker><length>,'."""
    idx = data.find(marker)
    if idx == -1:
        return -1
    comma = data.index(b",", idx)
    return comma + 1


def _deinterleave(raw: np.ndarray) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    """Split a (CARD_HEIGHT, ROW_STRIDE) array into Y, M, C channels."""
    y = raw[:, 0::3][:, :CARD_WIDTH]
    m = raw[:, 1::3][:, :CARD_WIDTH]
    c = raw[:, 2::3][:, :CARD_WIDTH]
    return y, m, c


# ---------------------------------------------------------------------------
# Decoder
# ---------------------------------------------------------------------------

def decode_prn(prn_path: str | Path) -> dict:
    """Decode a Saturn PRN file into a color image and K-layer mask.

    Supports both old format (color only, K inferred from YMC=0) and
    new format (separate SKB K-layer block).

    Returns a dict with:
        "color"  : PIL.Image  RGB  (YMC converted to RGB)
        "k_mask" : PIL.Image  L    (255 = K-resin pixel, 0 = not K)
        "width"  : int
        "height" : int
    """
    prn_path = Path(prn_path)
    with open(prn_path, "rb") as f:
        data = f.read()

    color_offset = _find_data_after_marker(data, SCB_MARKER)
    if color_offset == -1:
        raise ValueError("No ESC SCB= command found in file")
    color_bytes = data[color_offset : color_offset + PIXEL_DATA_SIZE]
    if len(color_bytes) < PIXEL_DATA_SIZE:
        raise ValueError(
            f"File too small: expected {PIXEL_DATA_SIZE} bytes of color data "
            f"at offset 0x{color_offset:X}, got {len(color_bytes)}"
        )

    raw_color = np.frombuffer(color_bytes, dtype=np.uint8).reshape(
        CARD_HEIGHT, ROW_STRIDE
    )
    y_chan, m_chan, c_chan = _deinterleave(raw_color)

    k_offset = _find_data_after_marker(data, SKB_MARKER)
    if k_offset != -1:
        k_bytes = data[k_offset : k_offset + PIXEL_DATA_SIZE]
        raw_k = np.frombuffer(k_bytes, dtype=np.uint8).reshape(
            CARD_HEIGHT, ROW_STRIDE
        )
        k_ch, _, _ = _deinterleave(raw_k)
        k_mask = k_ch == 0
    else:
        k_mask = (y_chan == 0) & (m_chan == 0) & (c_chan == 0)

    color_img = Image.fromarray(
        np.stack([c_chan, m_chan, y_chan], axis=2), mode="RGB"
    )
    k_img = Image.fromarray((k_mask.astype(np.uint8) * 255), mode="L")

    return {
        "color": color_img,
        "k_mask": k_img,
        "width": CARD_WIDTH,
        "height": CARD_HEIGHT,
    }


def save_decoded(prn_path: str | Path, output_dir: str | Path | None = None) -> None:
    """Decode a PRN and save color + K-layer PNGs next to the source file."""
    prn_path = Path(prn_path)
    result = decode_prn(prn_path)

    if output_dir is None:
        output_dir = prn_path.parent
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    stem = prn_path.stem

    color_path = output_dir / f"{stem}_COLOR.png"
    result["color"].save(color_path)
    print(f"Saved color image : {color_path}")

    k_path = output_dir / f"{stem}_K.png"
    result["k_mask"].save(k_path)

    k_count = np.count_nonzero(np.array(result["k_mask"]))
    total = CARD_WIDTH * CARD_HEIGHT
    print(f"Saved K-layer mask: {k_path}  ({k_count}/{total} pixels, {100*k_count/total:.1f}%)")


# ---------------------------------------------------------------------------
# Encoder
# ---------------------------------------------------------------------------

def _load_template(template_prn: str | Path) -> dict:
    """Extract the static blocks from a template PRN.

    Returns a dict with:
        "color_preamble" : bytes from ESC PDM through the comma after SCB=<size>,
        "k_preamble"     : bytes from CR ESC SKP through the comma after SKB=<size>,
                           (empty bytes if template has no K layer)
    """
    template_prn = Path(template_prn)
    with open(template_prn, "rb") as f:
        data = f.read()

    sop_end = data.index(b"\r", data.index(b"\x1bSOP=")) + 1
    color_pixel_start = _find_data_after_marker(data, SCB_MARKER)

    k_pixel_start = _find_data_after_marker(data, SKB_MARKER)
    if k_pixel_start != -1:
        color_pixel_end = color_pixel_start + PIXEL_DATA_SIZE
        k_preamble = data[color_pixel_end:k_pixel_start]
    else:
        k_preamble = b""

    return {
        "color_preamble": data[sop_end:color_pixel_start],
        "k_preamble": k_preamble,
    }


def _prepare_image(image_path: str | Path) -> np.ndarray:
    """Load an image and return a (CARD_HEIGHT, CARD_WIDTH, 3) uint8 RGB array.

    Handles RGBA by compositing onto white. Resizes to card dimensions.
    """
    img = Image.open(image_path)

    if img.mode == "RGBA":
        background = Image.new("RGB", img.size, (255, 255, 255))
        background.paste(img, mask=img.split()[3])
        img = background
    elif img.mode == "P":
        img = img.convert("RGBA")
        background = Image.new("RGB", img.size, (255, 255, 255))
        background.paste(img, mask=img.split()[3])
        img = background
    elif img.mode != "RGB":
        img = img.convert("RGB")

    img = img.resize((CARD_WIDTH, CARD_HEIGHT), Image.LANCZOS)
    return np.array(img, dtype=np.uint8)


def _interleave_ymc(y: np.ndarray, m: np.ndarray, c: np.ndarray) -> np.ndarray:
    """Pack Y, M, C channel arrays into a (CARD_HEIGHT, ROW_STRIDE) buffer."""
    out = np.full((CARD_HEIGHT, ROW_STRIDE), 0xFF, dtype=np.uint8)
    out[:, 0::3][:, :CARD_WIDTH] = y
    out[:, 1::3][:, :CARD_WIDTH] = m
    out[:, 2::3][:, :CARD_WIDTH] = c
    return out


def _rgb_to_layers(
    rgb: np.ndarray, k_threshold: int = 30
) -> tuple[np.ndarray, np.ndarray]:
    """Convert an RGB image to separate color and K-resin layers.

    Returns (color_data, k_data) each shaped (CARD_HEIGHT, ROW_STRIDE).
    Where K is active: color = 255 (no dye), K = 0 (apply resin).
    Where K is inactive: color = YMC values, K = 255.
    """
    y = rgb[:, :, 2].copy()  # B -> Y channel
    m = rgb[:, :, 1].copy()  # G -> M channel
    c = rgb[:, :, 0].copy()  # R -> C channel

    if k_threshold > 0:
        is_k = np.max(rgb, axis=2) < k_threshold
        y[is_k] = 255
        m[is_k] = 255
        c[is_k] = 255
    else:
        is_k = np.zeros((CARD_HEIGHT, CARD_WIDTH), dtype=bool)

    color_data = _interleave_ymc(y, m, c)

    k_val = np.full((CARD_HEIGHT, CARD_WIDTH), 255, dtype=np.uint8)
    k_val[is_k] = 0
    k_data = _interleave_ymc(k_val, k_val, k_val)

    return color_data, k_data


def build_prn(
    image_path: str | Path,
    output_path: str | Path,
    *,
    k_threshold: int = 30,
    template_prn: str | Path | None = None,
    job_id: int = 1,
) -> Path:
    """Encode a PNG/image file into a Saturn PRN file.

    Args:
        image_path:   Source image (PNG, JPEG, etc.).
        output_path:  Destination .prn file.
        k_threshold:  Pixels with max(R,G,B) below this value trigger K resin.
                      Set to 0 to disable K entirely.
        template_prn: Template PRN to extract the static header from.
        job_id:       Job ID number for the ESC commands.

    Returns:
        Path to the written PRN file.
    """
    if template_prn is None:
        template_prn = DEFAULT_TEMPLATE
    template_prn = Path(template_prn)
    output_path = Path(output_path)

    tpl = _load_template(template_prn)

    rgb = _prepare_image(image_path)
    color_data, k_data = _rgb_to_layers(rgb, k_threshold=k_threshold)

    k_ch = k_data[:, 0::3][:, :CARD_WIDTH]
    k_count = int(np.sum(k_ch == 0))

    header = (
        f"\x1bSOJ={job_id},Card_1-[Front]\r"
        f"\x1bSOP={job_id},1\r"
    ).encode("ascii")

    trailer = (
        f"\r\x1bEOP={job_id},1\r"
        f"\x1bEOJ={job_id},1\r"
    ).encode("ascii")

    with open(output_path, "wb") as f:
        f.write(header)
        f.write(tpl["color_preamble"])
        f.write(color_data.tobytes())
        if tpl["k_preamble"]:
            f.write(tpl["k_preamble"])
            f.write(k_data.tobytes())
        f.write(trailer)

    total = CARD_WIDTH * CARD_HEIGHT
    print(f"Encoded {image_path} -> {output_path}")
    print(f"  {CARD_WIDTH}x{CARD_HEIGHT}, K-resin pixels: {k_count}/{total} ({100*k_count/total:.1f}%)")

    return output_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _cli_decode(args: argparse.Namespace) -> None:
    for path in args.files:
        print(f"\n--- {path} ---")
        save_decoded(path)


def _cli_encode(args: argparse.Namespace) -> None:
    image_path = Path(args.image)
    if args.output:
        output_path = Path(args.output)
    else:
        output_path = image_path.with_suffix(".prn")

    k_threshold = 0 if args.no_k else args.k_threshold

    build_prn(
        image_path,
        output_path,
        k_threshold=k_threshold,
        job_id=args.job_id,
    )


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Saturn Card Printer PRN encoder/decoder"
    )
    sub = parser.add_subparsers(dest="command", required=True)

    dec = sub.add_parser("decode", help="Decode .prn files to PNG images")
    dec.add_argument("files", nargs="+", help="PRN file(s) to decode")

    enc = sub.add_parser("encode", help="Encode a PNG image to a .prn file")
    enc.add_argument("image", help="Source image (PNG, JPEG, etc.)")
    enc.add_argument("-o", "--output", help="Output .prn path (default: <image>.prn)")
    enc.add_argument(
        "-k", "--k-threshold", type=int, default=30,
        help="Pixels with max(R,G,B) below this trigger K resin (default: 30)",
    )
    enc.add_argument(
        "--no-k", action="store_true",
        help="Disable K-resin layer entirely (print black using YMC mix only)",
    )
    enc.add_argument(
        "--job-id", type=int, default=1,
        help="Job ID for ESC commands (default: 1)",
    )

    args = parser.parse_args()
    if args.command == "decode":
        _cli_decode(args)
    elif args.command == "encode":
        _cli_encode(args)


if __name__ == "__main__":
    main()
