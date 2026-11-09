from argparse import ArgumentParser
from re import fullmatch
from zipfile import ZipFile, ZIP_DEFLATED

parser = ArgumentParser(description="Package an editable PPTX with optional click-controlled fade transitions.")
parser.add_argument("source")
parser.add_argument("output")
parser.add_argument("--fade", action="store_true")
args = parser.parse_args()
transition = '<p:transition spd="med" advClick="1"><p:fade/></p:transition>'
with ZipFile(args.source) as source, ZipFile(args.output, "w", ZIP_DEFLATED) as output:
    for name in source.namelist():
        data = source.read(name)
        if args.fade and fullmatch(r"ppt/slides/slide[2-8]\.xml", name):
            data = data.replace(b"</p:sld>", (transition + "</p:sld>").encode(), 1)
        output.writestr(name, data)
