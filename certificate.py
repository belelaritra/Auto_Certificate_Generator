from PIL import Image, ImageDraw, ImageFont

# ==== COLOUR DICTIONARY WITH HEX COLOUR VALUE
colour_dict = {
    'Black': '#000000',
    'White': '#FFFFFF',
    'Red': '#FF0000',
    'Lime': '#00FF00',
    'Blue': '#FFFF00',
    'Yellow': '#FFFF00'
}
# =============================================== PREVIEW CERTIFICATE =================================================
def previewcertificate(Path, fontstyle, Colour, fontsize, coord):
# === CONVERT COORD STRING INTO INT TYPE (X , Y)
    splitcoord = coord.split(",")
    x = splitcoord[0]
    y = splitcoord[1]
    location = (int(x), int(y))
# === FONT STYLE & FONT SIZE
    font = ImageFont.truetype(fontstyle, int(fontsize))
# === OPEN TEMPLATE
    image = Image.open(Path)
    draw = ImageDraw.Draw(image)
# === EMBEDDED NAME
    draw.text(location, 'Sample Name', fill=colour_dict[Colour], font=font)
# === SAVE IMAGE FILE
    image.save('sample_image.jpeg')

# =============================================== CREATE CERTIFICATE =================================================
def createcertificate(Name, Path, fontstyle, Colour, fontsize, coord):
# === CONVERT COORD STRING INTO INT TYPE (X , Y)
    splitcoord = coord.split(",")
    x = splitcoord[0]
    y = splitcoord[1]
    location = (int(x), int(y))
# === FONT STYLE & FONT SIZE
    font = ImageFont.truetype(fontstyle, int(fontsize))
# === OPEN TEMPLATE
    image = Image.open(Path)
    draw = ImageDraw.Draw(image)
# === EMBEDDED NAME
    draw.text(location, Name, fill=colour_dict[Colour], font=font)
# === IMAGE NAME VAR (FIRSTNAME_LASTNAME.JPEG)
    image_name = Name.replace(' ', '_') + ".jpeg"
# === SAVE IMAGE FILE
    image.save(image_name)
# === RETURN IMAGE FILE NAME
    return image_name
