from PIL import Image, ImageDraw

def create_icon():
    """Creates a simple ICO file for the executable."""
    width, height = 64, 64
    color1, color2 = "black", "#4a85e8" # A blue color
    image = Image.new("RGB", (width, height), color1)
    dc = ImageDraw.Draw(image)
    # A simple 'L' shape for Launcher
    dc.rectangle((16, 16, 28, 48), fill=color2)
    dc.rectangle((16, 36, 48, 48), fill=color2)
    
    # ICO format requires specific sizes, let's provide a few
    icon_sizes = [(16,16), (32,32), (48,48), (64,64)]
    image.save("icon.ico", "ICO", sizes=icon_sizes)

if __name__ == "__main__":
    create_icon()
    print("icon.ico created successfully.")
