"""
Saturn Card Printer PRN decoder / encoder.

PRN file structure (reverse-engineered):
    ESC SOJ=<n>,Card_1-[Front] CR
    ESC SOP=<n>,1 CR
    ESC PDM=2512,<2512 bytes TFSM header>
    CR ESC SCP=80,<80 bytes color-panel config>
    CR ESC SCB=<size>,<pixel data>
    CR ESC EOP=<n>,1 CR
    ESC EOJ=<n>,1 CR

Pixel data is 8-bit-per-channel YMC, interleaved byte-by-byte:
    Y0 M0 C0  Y1 M1 C1  ...  Y1012 M1012 C1012  PAD
Row stride = 3040 bytes  (1013 * 3 + 1 padding 0xFF).
643 rows for a standard CR-80 card at 300 DPI.

Subtractive ink model: 255 = no ink (white), 0 = full ink.
K-resin trigger: when Y=M=C=0 exactly, the printer applies K resin
instead of mixing YMC dye.
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
DEFAULT_TEMPLATE = Path(__file__).parent / "files" / "normal.prn"


# ---------------------------------------------------------------------------
# Decoder
# ---------------------------------------------------------------------------

def _find_pixel_data(data: bytes) -> int:
    """Find the byte offset where pixel data begins (just after 'ESC SCB=<len>,')."""
    idx = data.find(SCB_MARKER)
    if idx == -1:
        raise ValueError("No ESC SCB= command found in file")
    comma = data.index(b",", idx)
    return comma + 1


def decode_prn(prn_path: str | Path) -> dict:
    """Decode a Saturn PRN file into a color image and K-layer mask.

    Returns a dict with:
        "color"  : PIL.Image  RGB  (YMC converted to RGB)
        "k_mask" : PIL.Image  L    (255 = K-resin pixel, 0 = not K)
        "width"  : int
        "height" : int
    """
    prn_path = Path(prn_path)
    with open(prn_path, "rb") as f:
        data = f.read()

    offset = _find_pixel_data(data)
    pixel_bytes = data[offset : offset + PIXEL_DATA_SIZE]
    if len(pixel_bytes) < PIXEL_DATA_SIZE:
        raise ValueError(
            f"File too small: expected {PIXEL_DATA_SIZE} bytes of pixel data "
            f"at offset 0x{offset:X}, got {len(pixel_bytes)}"
        )

    raw = np.frombuffer(pixel_bytes, dtype=np.uint8).reshape(CARD_HEIGHT, ROW_STRIDE)

    y_chan = raw[:, 0::3][:, :CARD_WIDTH]
    m_chan = raw[:, 1::3][:, :CARD_WIDTH]
    c_chan = raw[:, 2::3][:, :CARD_WIDTH]

    k_mask = (y_chan == 0) & (m_chan == 0) & (c_chan == 0)

    # Subtractive to RGB: Cyan→R, Magenta→G, Yellow→B
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

def _load_template(template_prn: str | Path) -> bytes:
    """Extract the static middle block from a template PRN.

    The middle block spans from ``ESC PDM=...`` through the comma after
    ``ESC SCB=1954720,``.  It is identical across print jobs regardless of
    job ID.
    """
    template_prn = Path(template_prn)
    with open(template_prn, "rb") as f:
        data = f.read()

    # Find the start of ESC PDM (first ESC after SOP)
    sop_end = data.index(b"\r", data.index(b"\x1bSOP=")) + 1
    pixel_start = _find_pixel_data(data)

    return data[sop_end:pixel_start]


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


def _rgb_to_ymc(rgb: np.ndarray, k_threshold: int = 30) -> np.ndarray:
    """Convert an RGB image array to interleaved YMC pixel data.

    Channel values map directly to RGB light intensity (no inversion needed):
      Y_channel = B,  M_channel = G,  C_channel = R
    (255 = no ink = full light, 0 = full ink = light absorbed)

    Pixels where max(R,G,B) < k_threshold are set to Y=M=C=0 (K-resin trigger).
    Returns a flat bytes-ready array of shape (CARD_HEIGHT, ROW_STRIDE).
    """
    y = rgb[:, :, 2].copy()  # B -> Y channel
    m = rgb[:, :, 1].copy()  # G -> M channel
    c = rgb[:, :, 0].copy()  # R -> C channel

    if k_threshold > 0:
        # K-resin: force Y=M=C=0 for near-black pixels
        is_k = np.max(rgb, axis=2) < k_threshold
        y[is_k] = 0
        m[is_k] = 0
        c[is_k] = 0
    else:
        # K disabled: clamp so Y=M=C never hits exact (0,0,0),
        # which would accidentally trigger K-resin.
        all_zero = (y == 0) & (m == 0) & (c == 0)
        y[all_zero] = 1
        m[all_zero] = 1
        c[all_zero] = 1

    # Interleave into row-stride format: Y0 M0 C0  Y1 M1 C1 ... PAD
    out = np.full((CARD_HEIGHT, ROW_STRIDE), 0xFF, dtype=np.uint8)
    out[:, 0::3][:, :CARD_WIDTH] = y
    out[:, 1::3][:, :CARD_WIDTH] = m
    out[:, 2::3][:, :CARD_WIDTH] = c

    return out


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
        template_prn: Template PRN to extract the static header from.
        job_id:       Job ID number for the ESC commands.

    Returns:
        Path to the written PRN file.
    """
    if template_prn is None:
        template_prn = DEFAULT_TEMPLATE
    template_prn = Path(template_prn)
    output_path = Path(output_path)

    middle_block = _load_template(template_prn)

    rgb = _prepare_image(image_path)
    pixel_data = _rgb_to_ymc(rgb, k_threshold=k_threshold)

    # Count K pixels from the interleaved data (skip padding byte per row)
    y_out = pixel_data[:, 0::3][:, :CARD_WIDTH]
    m_out = pixel_data[:, 1::3][:, :CARD_WIDTH]
    c_out = pixel_data[:, 2::3][:, :CARD_WIDTH]
    k_count = int(np.sum((y_out == 0) & (m_out == 0) & (c_out == 0)))

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
        f.write(middle_block)
        f.write(pixel_data.tobytes())
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
