"""
SVG Banner Renderer
==================

This module provides functionality to render SVG files as images
for use in Tkinter applications.

Author: Workflow Manager System
Version: 1.0.0
"""

import tkinter as tk
from PIL import Image, ImageTk
import io
import cairosvg
import os


class SVGBannerRenderer:
    """Renderer for SVG banner files"""
    
    def __init__(self, svg_file_path):
        """
        Initialize SVG banner renderer
        
        Args:
            svg_file_path: Path to the SVG file
        """
        self.svg_file_path = svg_file_path
        self.drawing = None
        self.image = None
        self.photo = None
        
    def render_svg(self, width=None, height=None):
        """
        Render SVG to PIL Image using cairosvg
        
        Args:
            width: Target width (optional)
            height: Target height (optional)
            
        Returns:
            PIL Image object
        """
        try:
            # Read SVG file
            with open(self.svg_file_path, 'rb') as svg_file:
                svg_data = svg_file.read()
            
            # Convert SVG to PNG using cairosvg
            if width and height:
                png_data = cairosvg.svg2png(
                    bytestring=svg_data,
                    output_width=width,
                    output_height=height,
                    background_color='transparent'
                )
            else:
                png_data = cairosvg.svg2png(bytestring=svg_data, background_color='transparent')
            
            # Convert to PIL Image with alpha preserved
            self.image = Image.open(io.BytesIO(png_data)).convert('RGBA')
            return self.image
            
        except Exception as e:
            print(f"Error rendering SVG with cairosvg: {e}")
            # Try to create banner with logos instead of simple fallback
            return self.create_banner_with_logos(width or 1920, height or 95)
    
    def create_fallback_image(self, width, height):
        """
        Create a fallback banner image when SVG rendering fails
        
        Args:
            width: Image width
            height: Image height
            
        Returns:
            PIL Image object
        """
        from PIL import ImageDraw, ImageFont
        
        # Create a simple banner image
        img = Image.new('RGB', (width, height), color='#b3b3b3')
        draw = ImageDraw.Draw(img)
        
        try:
            # Try to use a system font
            font = ImageFont.truetype("arial.ttf", 16)
        except:
            font = ImageFont.load_default()
        
        # Draw banner text
        text = "WORKFLOW MANAGER - Status: No Alarms | 2 Active Users | System Up"
        text_bbox = draw.textbbox((0, 0), text, font=font)
        text_width = text_bbox[2] - text_bbox[0]
        text_height = text_bbox[3] - text_bbox[1]
        
        # Center the text
        x = (width - text_width) // 2
        y = (height - text_height) // 2
        
        draw.text((x, y), text, fill='black', font=font)
        
        return img
    
    def create_banner_with_logos(self, width, height):
        """
        Create a banner image with Lumen and TXO logos
        
        Args:
            width: Image width
            height: Image height
            
        Returns:
            PIL Image object
        """
        from PIL import ImageDraw, ImageFont
        
        # Create banner image
        img = Image.new('RGB', (width, height), color='#b3b3b3')
        draw = ImageDraw.Draw(img)
        
        try:
            # Load and render Lumen logo
            lumen_img = self.render_logo_svg("Lumen.svg", 200, 50)
            if lumen_img:
                img.paste(lumen_img, (20, (height - 50) // 2), lumen_img if lumen_img.mode == 'RGBA' else None)
            
            # Load and render TXO logo
            txo_img = self.render_logo_svg("TXO.svg", 100, 50)
            if txo_img:
                img.paste(txo_img, (width - 120, (height - 50) // 2), txo_img if txo_img.mode == 'RGBA' else None)
            
            # Add status text in the center
            try:
                font = ImageFont.truetype("arial.ttf", 14)
            except:
                font = ImageFont.load_default()
            
            status_text = "No Status Alarms  |  No Database Alarms  |  2 Active Users  |  System up time: 4 days, 4 hours, 12 min"
            text_bbox = draw.textbbox((0, 0), status_text, font=font)
            text_width = text_bbox[2] - text_bbox[0]
            text_height = text_bbox[3] - text_bbox[1]
            
            # Center the text
            x = (width - text_width) // 2
            y = (height - text_height) // 2
            
            draw.text((x, y), status_text, fill='black', font=font)
            
        except Exception as e:
            print(f"Error creating banner with logos: {e}")
        
        return img
    
    def render_logo_svg(self, svg_file, target_width, target_height):
        """
        Render a logo SVG to PIL Image
        
        Args:
            svg_file: Path to SVG file
            target_width: Target width
            target_height: Target height
            
        Returns:
            PIL Image object or None
        """
        try:
            if not os.path.exists(svg_file):
                return None
                
            # Read SVG file
            with open(svg_file, 'rb') as f:
                svg_data = f.read()
            
            # Convert SVG to PNG using cairosvg
            png_data = cairosvg.svg2png(
                bytestring=svg_data,
                output_width=target_width,
                output_height=target_height,
                background_color='transparent'
            )
            
            # Convert to PIL Image with alpha preserved
            return Image.open(io.BytesIO(png_data)).convert('RGBA')
            
        except Exception as e:
            print(f"Error rendering logo {svg_file}: {e}")
            return None
    
    def get_tkinter_photo(self, width=None, height=None):
        """
        Get Tkinter PhotoImage from SVG
        
        Args:
            width: Target width (optional)
            height: Target height (optional)
            
        Returns:
            Tkinter PhotoImage object
        """
        if not self.image:
            self.render_svg(width, height)
        
        if self.image:
            # Convert PIL Image to Tkinter PhotoImage
            self.photo = ImageTk.PhotoImage(self.image)
            return self.photo
        
        return None
    
    def create_banner_widget(self, parent, width=1920, height=95):
        """
        Create a banner widget with the rendered SVG
        
        Args:
            parent: Parent Tkinter widget
            width: Banner width
            height: Banner height
            
        Returns:
            Tkinter Label widget with the banner image
        """
        # Render SVG to specific dimensions
        image = self.render_svg(width, height)
        
        if image:
            # Convert to Tkinter PhotoImage
            photo = ImageTk.PhotoImage(image)
            
            # Create label with the image
            banner_label = tk.Label(parent, image=photo, bg='#b3b3b3')
            
            # Store the photo reference in the label itself to prevent garbage collection
            banner_label.photo = photo
            
            # Configure the label to scale with the parent
            banner_label.pack_propagate(False)
            
            return banner_label
        else:
            # Fallback: create empty label
            return tk.Label(parent, text="Banner Error", bg='#b3b3b3')


def create_banner_from_svg(svg_file_path, parent, width=1920, height=95):
    """
    Convenience function to create a banner widget from SVG file
    
    Args:
        svg_file_path: Path to SVG file
        parent: Parent Tkinter widget
        width: Banner width
        height: Banner height
        
    Returns:
        Tkinter Label widget with the banner image
    """
    try:
        renderer = SVGBannerRenderer(svg_file_path)
        return renderer.create_banner_widget(parent, width, height)
    except Exception as e:
        print(f"Error in create_banner_from_svg: {e}")
        return None


if __name__ == "__main__":
    # Test the SVG renderer
    root = tk.Tk()
    root.title("SVG Banner Test")
    root.geometry("1920x200")
    
    # Create banner
    svg_path = "Workflow manager - login screen - Top Banner.svg"
    if os.path.exists(svg_path):
        banner = create_banner_from_svg(svg_path, root, width=1920, height=95)
        banner.pack(fill=tk.X, pady=10)
    else:
        print(f"SVG file not found: {svg_path}")
    
    root.mainloop()
