import numpy as np
from PIL import Image
import os

def process_saturn_prn(prn_path):
    # Hardware constants derived from your previous tests
    STRIDE = 1520
    OFFSET = 0xa60
    
    with open(prn_path, 'rb') as f:
        data = f.read()

    # Align 16-bit buffer
    pixel_data = data[OFFSET:]
    if len(pixel_data) % 2 != 0:
        pixel_data = pixel_data[:-1]

    raw = np.frombuffer(pixel_data, dtype='<H')
    height = len(raw) // STRIDE
    grid = raw[:height * STRIDE].reshape((height, STRIDE))

    # DATA-DRIVEN DETECTION: Measure vertical coherence (jitter)
    # If the period is correct, a column will represent a consistent 
    # physical position on the card, resulting in lower row-to-row variance.
    def measure_jitter(p):
        w = STRIDE // p
        # Using Magenta (index 1) as a stable reference channel
        sample = grid[:, 1::p][:, :w].astype(float)
        return np.mean(np.abs(sample[:-1, :] - sample[1:, :]))

    # Test both periods (3-column CMY vs 4-column CMYK)
    period = 3 if measure_jitter(3) < measure_jitter(4) else 4
    width = STRIDE // period

    # Channel Extraction
    c_raw = grid[:, 0::period][:, :width]
    m_raw = grid[:, 1::period][:, :width]
    y_raw = grid[:, 2::period][:, :width]
    
    # Apply your confirmed -1 Cyan shift for horizontal alignment
    c_aligned = np.roll(c_raw, -1, axis=1)

    # Subtractive logic conversion (0xFFFF = White card, 0x0000 = Ink)
    r = (c_aligned.astype(float) / 65535.0) * 255
    g = (m_raw.astype(float) / 65535.0) * 255
    b = (y_raw.astype(float) / 65535.0) * 255

    # Color Merging
    rgb_stack = np.stack([r, g, b], axis=2).astype(np.uint8)
    color_img = Image.fromarray(rgb_stack, mode='RGB')
    
    # Scale to standard CR-80 card ratio (approx 1.58)
    final_h = int(width / 1.58)
    color_img.resize((width, final_h), Image.LANCZOS).save("COLOR_OUTPUT.png")

    # K-Layer Logic: Only for 4-column mode and only if data is present
    if period == 4:
        k_channel = grid[:, 3::4][:, :width]
        # Content Check: Detect actual ink (pixels significantly darker than 0xFFFF)
        if np.any(k_channel < 65000):
            k_intensity = (k_channel.astype(float) / 65535.0) * 255
            k_img = Image.fromarray(k_intensity.astype(np.uint8), mode='L')
            k_img.resize((width, final_h), Image.LANCZOS).save("K_LAYER_OUTPUT.png")

if __name__ == "__main__":
    process_saturn_prn("center-pixels.prn")
