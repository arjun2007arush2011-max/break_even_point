#!/usr/bin/env python3
"""
Generates high-resolution PNG and SVG QR codes for:
1. Live GitHub Pages Web App: https://arjun2007arush2011-max.github.io/break_even_point/
2. GitHub Repository: https://github.com/arjun2007arush2011-max/break_even_point
"""

import sys
import os
import zlib
import struct

user_site = os.path.expanduser('~/Library/Python/3.9/lib/python/site-packages')
if user_site not in sys.path:
    sys.path.insert(0, user_site)

from reportlab.graphics.barcode import qr
from reportlab.graphics.shapes import Drawing
from reportlab.graphics import renderSVG

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
URL_LIVE = "https://arjun2007arush2011-max.github.io/break_even_point/"
URL_GITHUB = "https://github.com/arjun2007arush2011-max/break_even_point"


def generate_qr_png(data_url, output_path, scale=16, border=4):
    """Generates a crisp, lossless grayscale PNG QR code with pure Python."""
    q = qr.QrCodeWidget(data_url)
    q.getBounds()
    modules = q.qr.modules
    n = len(modules)
    
    grid_size = n + 2 * border
    width = grid_size * scale
    height = grid_size * scale
    
    raw_rows = []
    for y in range(grid_size):
        mod_y = y - border
        row_bytes = bytearray()
        for x in range(grid_size):
            mod_x = x - border
            if 0 <= mod_y < n and 0 <= mod_x < n and modules[mod_y][mod_x]:
                pixel = 0    # Black module
            else:
                pixel = 255  # White background
            row_bytes.extend([pixel] * scale)
        for _ in range(scale):
            raw_rows.append(b'\x00' + bytes(row_bytes))
            
    raw_data = b''.join(raw_rows)
    compressed = zlib.compress(raw_data, 9)
    
    def chunk(chunk_type, chunk_data):
        c = chunk_type + chunk_data
        crc = struct.pack('>I', zlib.crc32(c) & 0xffffffff)
        return struct.pack('>I', len(chunk_data)) + c + crc
        
    png_header = b'\x89PNG\r\n\x1a\n'
    ihdr = chunk(b'IHDR', struct.pack('>IIBBBBB', width, height, 8, 0, 0, 0, 0))
    idat = chunk(b'IDAT', compressed)
    iend = chunk(b'IEND', b'')
    
    with open(output_path, 'wb') as f:
        f.write(png_header + ihdr + idat + iend)
    print(f"[OK] Generated PNG ({width}x{height}): {output_path}")


def generate_qr_svg(data_url, output_path, size=300):
    """Generates a scalable vector SVG QR code."""
    qr_code = qr.QrCodeWidget(data_url)
    bounds = qr_code.getBounds()
    w = bounds[2] - bounds[0]
    h = bounds[3] - bounds[1]
    d = Drawing(size, size, transform=[float(size)/w, 0, 0, float(size)/h, 0, 0])
    d.add(qr_code)
    renderSVG.drawToFile(d, output_path)
    print(f"[OK] Generated SVG ({size}x{size}): {output_path}")


if __name__ == "__main__":
    png_live = os.path.join(BASE_DIR, "qrcode_live.png")
    svg_live = os.path.join(BASE_DIR, "qrcode_live.svg")
    png_github = os.path.join(BASE_DIR, "qrcode_github.png")
    svg_github = os.path.join(BASE_DIR, "qrcode_github.svg")

    generate_qr_png(URL_LIVE, png_live)
    generate_qr_svg(URL_LIVE, svg_live)

    generate_qr_png(URL_GITHUB, png_github)
    generate_qr_svg(URL_GITHUB, svg_github)
