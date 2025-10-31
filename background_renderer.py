import tkinter as tk
from PIL import Image, ImageTk
import io
import os

class BackgroundRenderer:
    """Renderer for application background image files (PNG preferred, SVG supported)"""
    
    def __init__(self, image_file_path="background.png"):
        """
        Initialize background renderer
        Args:
            image_file_path: Path to the background image file (PNG/JPG/SVG)
        """
        self.image_file_path = image_file_path
        self.image = None
        self.photo = None
    
    def render_background(self, width=1920, height=1080):
        """
        Render background image to a PIL Image.
        - If file is PNG/JPG, load and resize.
        - If file is SVG, convert to PNG via cairosvg (if available).
        Args:
            width: Target width (default 1920)
            height: Target height (default 1080)
        Returns:
            PIL Image object
        """
        try:
            # Check if file exists
            if not os.path.exists(self.image_file_path):
                print(f"Background image file not found: {self.image_file_path}")
                return self.create_fallback_background(width, height)
            
            _, ext = os.path.splitext(self.image_file_path.lower())

            if ext in [".png", ".jpg", ".jpeg", ".webp", ".bmp"]:
                # Load raster image and resize to fit
                img = Image.open(self.image_file_path).convert("RGB")
                self.image = img.resize((width, height), Image.LANCZOS)
                return self.image
            elif ext == ".svg":
                # Lazy import cairosvg only when needed
                try:
                    import cairosvg  # type: ignore
                except Exception as import_error:
                    print(f"cairosvg not available for SVG conversion: {import_error}")
                    return self.create_fallback_background(width, height)

                with open(self.image_file_path, 'rb') as svg_file:
                    svg_data = svg_file.read()

                png_data = cairosvg.svg2png(
                    bytestring=svg_data,
                    output_width=width,
                    output_height=height
                )

                self.image = Image.open(io.BytesIO(png_data))
                return self.image
            else:
                print(f"Unsupported background format: {ext}")
                return self.create_fallback_background(width, height)
            
        except Exception as e:
            print(f"Error rendering background image: {e}")
            return self.create_fallback_background(width, height)
    
    def create_fallback_background(self, width, height):
        """
        Create a fallback background when SVG rendering fails
        Args:
            width: Image width
            height: Image height
        Returns:
            PIL Image object with gradient background
        """
        from PIL import ImageDraw
        
        # Create a gradient background similar to the SVG
        img = Image.new('RGB', (width, height), color='#1a1a1a')
        draw = ImageDraw.Draw(img)
        
        # Create a simple radial gradient effect
        center_x, center_y = width // 2, height // 2
        max_radius = max(width, height) // 2
        
        for radius in range(max_radius, 0, -1):
            # Calculate color based on radius (darker towards edges)
            intensity = int(255 * (1 - radius / max_radius))
            color = (intensity // 3, intensity // 2, intensity // 2)  # Blue-gray gradient
            
            # Draw circle
            left = center_x - radius
            top = center_y - radius
            right = center_x + radius
            bottom = center_y + radius
            
            draw.ellipse([left, top, right, bottom], fill=color)
        
        return img
    
    def get_tkinter_photo(self, width=1920, height=1080):
        """
        Get Tkinter PhotoImage from background
        """
        if not self.image:
            self.render_background(width, height)
        
        if self.image:
            self.photo = ImageTk.PhotoImage(self.image)
            return self.photo
        return None
    
    def create_background_widget(self, parent, width=1920, height=1080):
        """
        Create a background widget with the rendered background
        Args:
            parent: Parent Tkinter widget
            width: Background width
            height: Background height
        Returns:
            Tkinter Label widget with the background image
        """
        image = self.render_background(width, height)
        
        if image:
            photo = ImageTk.PhotoImage(image)
            background_label = tk.Label(parent, image=photo, highlightthickness=0)
            background_label.image = photo  # Store the photo reference
            background_label.place(x=0, y=0, relwidth=1, relheight=1)
            return background_label
        else:
            return tk.Label(parent, text="Background Error", bg='#1a1a1a')
    
    def apply_background_to_window(self, window, width=None, height=None):
        """
        Apply background to an entire window
        Args:
            window: Tkinter window/widget
            width: Window width (optional, will detect if not provided)
            height: Window height (optional, will detect if not provided)
        """
        if width is None:
            width = window.winfo_screenwidth()
        if height is None:
            height = window.winfo_screenheight()
        
        # Create background widget
        background_widget = self.create_background_widget(window, width, height)
        
        # Send to back so other widgets appear on top
        background_widget.lower()
        
        # Store reference to prevent garbage collection
        window._background_widget = background_widget
        
        return background_widget

def create_background_from_file(image_file_path, parent, width=1920, height=1080):
    """
    Convenience function to create a background widget from an image file
    """
    try:
        renderer = BackgroundRenderer(image_file_path)
        return renderer.create_background_widget(parent, width, height)
    except Exception as e:
        print(f"Error in create_background_from_file: {e}")
        return None
