"""Generate the application icon (icon.ico) for Automatic Writing Assistant."""
import struct
import io
import sys
import math


def generate_ico_pure_python(path):
    """Generate a multi-size ICO file with a pen/quill design using pure Python."""
    sizes = [16, 32, 48, 64, 128, 256]
    images = []

    for size in sizes:
        pixels = _draw_icon(size)
        png_data = _encode_bmp_rgba(pixels, size)
        images.append((size, png_data))

    _write_ico(path, images)
    print(f"Icon saved to: {path}")


def _draw_icon(size):
    """Draw the icon at a given size. Returns list of (r, g, b, a) tuples."""
    pixels = [(0, 0, 0, 0)] * (size * size)

    # Colors matching the app theme
    bg_r, bg_g, bg_b = 0x0d, 0x0d, 0x1a       # Dark background
    accent_r, accent_g, accent_b = 0x6c, 0x63, 0xff  # Purple accent
    glow_r, glow_g, glow_b = 0x7c, 0x4d, 0xff  # Glow color
    white_r, white_g, white_b = 0xe0, 0xe0, 0xf0  # Text white

    cx, cy = size / 2.0, size / 2.0
    radius = size / 2.0 - 1

    for y in range(size):
        for x in range(size):
            px, py = x + 0.5, y + 0.5
            dx, dy = px - cx, py - cy
            dist = math.sqrt(dx * dx + dy * dy)

            if dist <= radius:
                # Rounded square background with gradient
                edge = max(abs(dx), abs(dy))
                corner_dist = math.sqrt(dx * dx + dy * dy)
                rr = radius * 0.85
                inside_sq = edge < rr

                if inside_sq or dist <= radius:
                    # Gradient background (darker at edges)
                    t = dist / radius
                    r = int(bg_r * (1 - t * 0.3))
                    g = int(bg_g * (1 - t * 0.3))
                    b = int(bg_b * (1 - t * 0.3))
                    a = 255

                    # Anti-alias edge
                    if dist > radius - 1.5:
                        a = int(max(0, min(255, (radius - dist) * 170)))

                    # Draw a stylized "W" letter in the center
                    # Normalize coordinates to 0-1 range relative to center
                    nx = (px - cx) / (size * 0.35)
                    ny = (py - cy) / (size * 0.35)

                    # "W" shape using line segments
                    w_hit = False
                    stroke = 2.5 / size * 7  # Adaptive stroke width

                    # W is made of 4 line segments
                    # Left down-stroke: (-0.8, -0.6) to (-0.4, 0.6)
                    w_hit = w_hit or _point_on_line(nx, ny,
                        -0.8, -0.6, -0.4, 0.6, stroke)
                    # Left-center up-stroke: (-0.4, 0.6) to (0.0, -0.1)
                    w_hit = w_hit or _point_on_line(nx, ny,
                        -0.4, 0.6, 0.0, -0.1, stroke)
                    # Right-center down-stroke: (0.0, -0.1) to (0.4, 0.6)
                    w_hit = w_hit or _point_on_line(nx, ny,
                        0.0, -0.1, 0.4, 0.6, stroke)
                    # Right up-stroke: (0.4, 0.6) to (0.8, -0.6)
                    w_hit = w_hit or _point_on_line(nx, ny,
                        0.4, 0.6, 0.8, -0.6, stroke)

                    if w_hit:
                        r, g, b = accent_r, accent_g, accent_b
                        a = 255

                    # Draw a subtle underline/cursor below the W
                    if -0.7 < nx < 0.7 and 0.75 < ny < 0.85:
                        r, g, b = glow_r, glow_g, glow_b
                        a = 255

                    # Small dots at top corners for decoration
                    for dot_x, dot_y in [(-0.6, -0.8), (0.6, -0.8)]:
                        dd = math.sqrt((nx - dot_x)**2 + (ny - dot_y)**2)
                        if dd < 0.12:
                            r, g, b = accent_r, accent_g, accent_b
                            a = int(min(255, (0.12 - dd) / 0.12 * 255))

                    pixels[y * size + x] = (r, g, b, a)
                else:
                    pixels[y * size + x] = (0, 0, 0, 0)
            else:
                pixels[y * size + x] = (0, 0, 0, 0)

    return pixels


def _point_on_line(px, py, x1, y1, x2, y2, thickness):
    """Check if point (px,py) is within thickness of line segment."""
    dx, dy = x2 - x1, y2 - y1
    len_sq = dx * dx + dy * dy
    if len_sq == 0:
        return math.sqrt((px - x1)**2 + (py - y1)**2) < thickness

    t = max(0, min(1, ((px - x1) * dx + (py - y1) * dy) / len_sq))
    proj_x = x1 + t * dx
    proj_y = y1 + t * dy
    dist = math.sqrt((px - proj_x)**2 + (py - proj_y)**2)
    return dist < thickness


def _encode_bmp_rgba(pixels, size):
    """Encode pixels as a BMP DIB (BITMAPINFOHEADER + pixel data) for ICO."""
    # BITMAPINFOHEADER (40 bytes)
    # Height is doubled in ICO (includes mask)
    header = struct.pack('<IiiHHIIiiII',
        40,          # biSize
        size,        # biWidth
        size * 2,    # biHeight (doubled for ICO)
        1,           # biPlanes
        32,          # biBitCount (RGBA)
        0,           # biCompression (BI_RGB)
        0,           # biSizeImage
        0,           # biXPelsPerMeter
        0,           # biYPelsPerMeter
        0,           # biClrUsed
        0,           # biClrImportant
    )

    # Pixel data (bottom-up, BGRA)
    pixel_data = bytearray()
    for y in range(size - 1, -1, -1):
        for x in range(size):
            r, g, b, a = pixels[y * size + x]
            pixel_data.extend([b, g, r, a])

    # AND mask (1bpp, bottom-up) - all zeros since we have alpha
    mask_row_bytes = ((size + 31) // 32) * 4
    mask_data = bytes(mask_row_bytes * size)

    return header + bytes(pixel_data) + mask_data


def _write_ico(path, images):
    """Write ICO file from list of (size, bmp_data) tuples."""
    num = len(images)

    # ICONDIR: reserved(2) + type(2) + count(2)
    header = struct.pack('<HHH', 0, 1, num)

    # Calculate offsets
    entries_size = 6 + num * 16  # Header + all ICONDIRENTRY
    offset = entries_size

    entries = []
    for size, data in images:
        w = size if size < 256 else 0
        h = size if size < 256 else 0
        entry = struct.pack('<BBBBHHII',
            w,           # bWidth
            h,           # bHeight
            0,           # bColorCount
            0,           # bReserved
            1,           # wPlanes
            32,          # wBitCount
            len(data),   # dwBytesInRes
            offset,      # dwImageOffset
        )
        entries.append(entry)
        offset += len(data)

    with open(path, 'wb') as f:
        f.write(header)
        for entry in entries:
            f.write(entry)
        for size, data in images:
            f.write(data)


def _encode_png(pixels, size):
    """Encode RGBA pixels as a minimal PNG file (pure Python, no deps)."""
    import zlib
    def _chunk(ctype, data):
        c = ctype + data
        crc = struct.pack('>I', zlib.crc32(c) & 0xFFFFFFFF)
        return struct.pack('>I', len(data)) + c + crc

    raw = bytearray()
    for y in range(size):
        raw.append(0)  # filter: None
        for x in range(size):
            r, g, b, a = pixels[y * size + x]
            raw.extend([r, g, b, a])

    sig = b'\x89PNG\r\n\x1a\n'
    ihdr = struct.pack('>IIBBBBB', size, size, 8, 6, 0, 0, 0)
    return sig + _chunk(b'IHDR', ihdr) + _chunk(b'IDAT', zlib.compress(bytes(raw), 9)) + _chunk(b'IEND', b'')


def generate_png(path, size=256):
    """Generate a PNG icon at the given size."""
    pixels = _draw_icon(size)
    data = _encode_png(pixels, size)
    with open(path, 'wb') as f:
        f.write(data)
    print(f"PNG icon saved to: {path}")


if __name__ == '__main__':
    out = 'icon.ico'
    if len(sys.argv) > 1:
        out = sys.argv[1]
    generate_ico_pure_python(out)
    # Also generate PNG for macOS/Linux
    png_path = out.rsplit('.', 1)[0] + '.png'
    generate_png(png_path, 256)
