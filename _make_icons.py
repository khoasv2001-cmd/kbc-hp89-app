"""Tăng kích cỡ logo trong PWA icons mà KHÔNG cắt phần nào.
Crop bỏ vùng trống xung quanh logo gốc -> scale to fit 88-92% canvas.
"""
from PIL import Image
import os

OUT_DIR = os.path.join(os.path.dirname(__file__), 'static')
SRC = os.path.join(OUT_DIR, 'kbc-logo.png')

RED = (198, 28, 45, 255)


def crop_to_content(img):
    """Cắt bỏ vùng trong suốt xung quanh, chỉ giữ nội dung logo."""
    # Lấy bounding box của pixel có alpha > 0
    bbox = img.getbbox()
    if bbox:
        return img.crop(bbox)
    return img


def fit_logo_to_canvas(logo, canvas_size, fill_ratio=0.90, bg=None):
    """Đặt logo vào giữa canvas vuông, scale sao cho cạnh dài nhất = fill_ratio * canvas.
    bg=None -> nền trong suốt; bg=tuple -> nền màu đặc."""
    canvas = Image.new('RGBA', (canvas_size, canvas_size), bg if bg else (0, 0, 0, 0))

    # Tính kích thước mới giữ tỷ lệ
    lw, lh = logo.size
    target = int(canvas_size * fill_ratio)
    if lw >= lh:
        new_w = target
        new_h = int(lh * target / lw)
    else:
        new_h = target
        new_w = int(lw * target / lh)

    logo_resized = logo.resize((new_w, new_h), Image.LANCZOS)
    # Căn giữa
    paste_x = (canvas_size - new_w) // 2
    paste_y = (canvas_size - new_h) // 2
    canvas.paste(logo_resized, (paste_x, paste_y), logo_resized)
    return canvas


def main():
    print(f'Đọc logo gốc: {SRC}')
    src = Image.open(SRC).convert('RGBA')
    print(f'  Kích thước gốc: {src.size}')

    # Crop bỏ vùng trống
    cropped = crop_to_content(src)
    print(f'  Sau khi crop: {cropped.size}')

    # KHÔNG sửa kbc-logo.png — chỉ regenerate icon files với logo lớn hơn

    # icon-192: nền trong suốt, logo chiếm 90% canvas
    icon192 = fit_logo_to_canvas(cropped, 192, fill_ratio=0.90)
    icon192.save(os.path.join(OUT_DIR, 'icon-192.png'), 'PNG')
    print('  icon-192.png: logo 90% canvas, transparent bg')

    # icon-512: tương tự
    icon512 = fit_logo_to_canvas(cropped, 512, fill_ratio=0.90)
    icon512.save(os.path.join(OUT_DIR, 'icon-512.png'), 'PNG')
    print('  icon-512.png: logo 90% canvas, transparent bg')

    # icon-maskable-512: Android maskable cần safe-zone, nền đỏ + logo 70% center
    mask = fit_logo_to_canvas(cropped, 512, fill_ratio=0.70, bg=RED)
    mask.save(os.path.join(OUT_DIR, 'icon-maskable-512.png'), 'PNG')
    print('  icon-maskable-512.png: logo 70% (safe zone), red bg')

    # favicon: 96x96, transparent bg
    fav = fit_logo_to_canvas(cropped, 96, fill_ratio=0.92)
    fav.save(os.path.join(OUT_DIR, 'favicon.png'), 'PNG')
    print('  favicon.png: logo 92% canvas')

    print('Done.')


if __name__ == '__main__':
    main()
