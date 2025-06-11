import os

ICON_COLOR = "rgb(50, 50, 200)"

ROOT = os.path.dirname(os.path.abspath(__file__)).replace("\\", "/") + "/icons"
DEST = "/".join(ROOT.split("/")[:-1]) + "/icons_colored"


for e in os.listdir(ROOT):
    file = open(ROOT + "/" + e, "r")
    text = file.read().replace("<path", f'<path fill="{ICON_COLOR}" ')
    text = text.replace("<rect", f'<rect fill="{ICON_COLOR}"')
    file.close()

    out_file = open(DEST + "/" + e, "w")
    out_file.write(text)
    out_file.close()
