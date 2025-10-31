from PIL import Image, ImageDraw, ImageFont, ImageOps
import textwrap

def scale_to_fit(img, max_width, max_height, allow_upscale=False):
    """Fit image inside (max_width, max_height) while preserving aspect ratio.
    Uses ImageOps.contain for robust, non-distorting scaling. Never upscales unless allowed.
    """
    if not allow_upscale and img.width <= max_width and img.height <= max_height:
        return img
    return ImageOps.contain(img, (int(max_width), int(max_height)), Image.LANCZOS)

def wrap_text(text, font, draw, max_width):
    """Wrap text to fit into a given pixel width."""
    wrapped_lines = []
    for paragraph in text.splitlines():
        if not paragraph:
            wrapped_lines.append("")  # preserve blank lines
            continue
        words = paragraph.split()
        line = ""
        for word in words:
            test_line = f"{line} {word}".strip()
            if draw.textlength(test_line, font=font) <= max_width:
                line = test_line
            else:
                wrapped_lines.append(line)
                line = word
        if line:
            wrapped_lines.append(line)
    return wrapped_lines

# --- Load background ---
background = Image.open(r"C:\Lumen\Workflow Manager\background.png").convert("RGBA")

# --- Load overlays ---
overlay_mid = Image.open(r"C:\Lumen\Workflow Manager\MidBanner2.png").convert("RGBA")
overlay_top = Image.open(r"C:\Lumen\Workflow Manager\TopBanner2.png").convert("RGBA")

# --- Scale overlays ---
# Keep original aspect; only shrink if too large (never upscale)
overlay_mid = scale_to_fit(
    overlay_mid,
    int(background.width * 0.7),   # allow up to 80% of width
    background.height // 3,
    allow_upscale=False
)
overlay_top = scale_to_fit(
    overlay_top,
    int(background.width * 0.9),
    overlay_top.height,
    allow_upscale=False
)

# --- Positioning ---
x_mid = (background.width - overlay_mid.width) // 2
y_mid = (background.height - overlay_mid.height) // 4
x_top = (background.width - overlay_top.width) // 2
y_top = 10  # offset 10px from the top edge

# --- Composite overlays ---
background.paste(overlay_mid, (x_mid, y_mid), overlay_mid)
background.paste(overlay_top, (x_top, y_top), overlay_top)

# --- Titles for login entries below MidBanner2 ---
draw = ImageDraw.Draw(background)
font_path = r"C:\Windows\Fonts\segoeui.ttf"
title_font = ImageFont.truetype(font_path, 25)

# Position 20px below the mid banner bottom
title_y = y_mid + overlay_mid.height + 80

def draw_centered(text: str, center_x: int, y: int):
    bbox = draw.textbbox((0, 0), text, font=title_font)
    text_w = bbox[2] - bbox[0]
    # Mild grey color for titles
    draw.text((center_x - text_w // 2, y), text, font=title_font, fill=(179, 179, 179, 255))

# Centers aligned with the entry placements in the GUI (relx 0.43 and 0.57)
username_cx = int(background.width * 0.43)
password_cx = int(background.width * 0.57)
draw_centered("Username", username_cx, title_y)
draw_centered("Password", password_cx, title_y)

# --- Disclaimer text ---
text = (
    "Authorized Use Only!\n\n"
    "The information presented through the use of this application is the property of this organization "
    "and is protected by intellectual property rights, and remains confidential to the organisation.\n\n"
    "By logging in you agree to keep this information confidential.\n\n"
    "You must be assigned an account on this system to access information and are only allowed "
    "to access information defined by the system administrators.\n\n"
    "Your activities will be monitored."
)

# Place disclaimer text starting at one-third of the screen width (left side)
left_margin = background.width // 3
right_margin = background.width // 3
bottom_margin = 60
max_width = max(50, background.width - left_margin - right_margin)
text_x = left_margin
text_y = int(background.height * (2/3))

# --- Fixed-size disclaimer text (do not scale with resolution) ---
draw = ImageDraw.Draw(background)
font_size = 14  # fixed size
font = ImageFont.truetype(font_path, font_size)
wrapped_lines = wrap_text(text, font, draw, max_width)

# Ensure we don't overflow the bottom; if overflow, trim lines
line_height = font.getbbox("Ay")[3] - font.getbbox("Ay")[1]
max_lines = max(1, (background.height - bottom_margin - text_y) // (line_height + 6))
wrapped_lines = wrapped_lines[:max_lines]

# Draw text
draw.multiline_text(
    (text_x, text_y),
    "\n".join(wrapped_lines),
    font=font,
    fill=(255, 255, 255, 255),
    spacing=6,
    align="left"
)

# Save result
background.save(r"C:\Lumen\Workflow Manager\combined.png", format="PNG")
