import numpy as np
from PIL import Image
import os

def decode_saturn_prn(prn_path, output_file="final_reconstruction.png"):
    # Hardware constants identified during reverse engineering
    WIDTH_STRIDE = 1520
    DATA_START = 0xa60
    MAX_COLS = WIDTH_STRIDE // 3  # 506 pixels per row
    
    with open(prn_path, 'rb') as f:
        data = f.read()

    # Ensure buffer is even for 16-bit unpacking
    pixel_data = data[DATA_START:]
    if len(pixel_data) % 2 != 0:
        pixel_data = pixel_data[:-1]

    # Load into 16-bit Little Endian grid
    raw_units = np.frombuffer(pixel_data, dtype='<H')
    height = len(raw_units) // WIDTH_STRIDE
    grid = raw_units[:height * WIDTH_STRIDE].reshape((height, WIDTH_STRIDE))

    # Column Mapping: Cyan=0, Magenta=1, Yellow=2
    # Truncate to max_cols to ensure matching shapes for stacking
    c_raw = grid[:, 0::3][:, :MAX_COLS]
    m_raw = grid[:, 1::3][:, :MAX_COLS]
    y_raw = grid[:, 2::3][:, :MAX_COLS]

    # Apply the 'Shift -1' correction to fix the Red/Cyan fringing
    # This aligns the Cyan sub-pixel with its M and Y partners
    c_aligned = np.roll(c_raw, -1, axis=1)

    # Subtractive logic: 0xFFFF (White/No Ink) -> 255 (Full Light)
    # 0x0000 (Black/Full Ink) -> 0 (No Light)
    r = (c_aligned.astype(float) / 65535.0) * 255
    g = (m_raw.astype(float) / 65535.0) * 255
    b = (y_raw.astype(float) / 65535.0) * 255

    # Merge into final RGB stack
    rgb = np.stack([r, g, b], axis=2).astype(np.uint8)
    
    img = Image.fromarray(rgb)

    # Aspect ratio correction for standard CR-80 card (approx 1.58:1)
    target_h = int(MAX_COLS / 1.58)
    img_final = img.resize((MAX_COLS, target_h), Image.LANCZOS)
    
    img_final.save(output_file)
    print(f"Successfully decoded {prn_path} to {output_file}")
    print(f"Dimensions: {MAX_COLS}x{height} (Scaled to {MAX_COLS}x{target_h})")

if __name__ == "__main__":
    # Change 'all.prn' to any Saturn PRN file
    decode_saturn_prn("black.prn")
