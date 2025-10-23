"""
Image utilities for quote generation.

This module provides functions for generating placeholder images
and loading product images for the quote document.
"""

from io import BytesIO
from PIL import Image, ImageDraw, ImageFont
import random

from constants import (
    DEFAULT_FONT_PATH,
    DEFAULT_FONT_SIZE,
    IMAGE_FORMAT
)


def generate_placeholder_image(width, height, seed=None, text=""):
    """
    Generate a placeholder image locally using PIL/Pillow.

    This function creates colorful placeholder images for testing.
    In production, you can replace this with:
    - Actual product images from your database/storage
    - Images from URLs (e.g., product.image_url)
    - Images from local file paths
    - External services like https://picsum.photos/200/300

    Example of using a real image from a URL:
        response = requests.get(product.image_url)
        return BytesIO(response.content)

    Args:
        width: Image width in pixels
        height: Image height in pixels
        seed: Optional seed for consistent colors
        text: Text to display on the image

    Returns:
        BytesIO object containing the image data
    """
    # Set random seed for consistent colors
    if seed is not None:
        random.seed(seed)

    # Generate a random pastel color
    r = random.randint(100, 255)
    g = random.randint(100, 255)
    b = random.randint(100, 255)
    bg_color = (r, g, b)

    # Create image
    img = Image.new('RGB', (width, height), color=bg_color)
    draw = ImageDraw.Draw(img)

    # Add text if provided
    if text:
        font = _load_font()
        _draw_text_with_shadow(draw, text, font, width, height)

    # Save to BytesIO
    img_io = BytesIO()
    img.save(img_io, format=IMAGE_FORMAT)
    img_io.seek(0)

    return img_io


def _load_font():
    """
    Load a font for text rendering.

    Returns:
        ImageFont object
    """
    try:
        return ImageFont.truetype(DEFAULT_FONT_PATH, DEFAULT_FONT_SIZE)
    except Exception:
        return ImageFont.load_default()


def _draw_text_with_shadow(draw, text, font, width, height):
    """
    Draw text centered on image with a shadow effect.

    Args:
        draw: ImageDraw object
        text: Text to draw
        font: Font to use
        width: Image width
        height: Image height
    """
    # Calculate text position (centered)
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    position = (
        (width - text_width) // 2,
        (height - text_height) // 2
    )

    # Draw text with shadow for better visibility
    shadow_color = (50, 50, 50)
    text_color = (255, 255, 255)

    draw.text(
        (position[0] + 2, position[1] + 2),
        text,
        fill=shadow_color,
        font=font
    )
    draw.text(position, text, fill=text_color, font=font)
