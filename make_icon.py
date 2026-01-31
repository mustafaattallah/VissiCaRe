from PIL import Image

Image.MAX_IMAGE_PIXELS = None

img = Image.open("icon maker.jpg").convert("RGBA")

# crop transparent border (optional)
bbox = img.getbbox()
if bbox:
    img = img.crop(bbox)

# shrink huge images to a reasonable working size (keeps quality)
MAX = 2048
w, h = img.size
if max(w, h) > MAX:
    scale = MAX / max(w, h)
    img = img.resize((int(w*scale), int(h*scale)), Image.Resampling.LANCZOS)

sizes = [256, 128, 64, 32, 16]
img.save("vissicare_icon.ico", sizes=[(s, s) for s in sizes])
print("Done! Created vissicare_icon.ico")
