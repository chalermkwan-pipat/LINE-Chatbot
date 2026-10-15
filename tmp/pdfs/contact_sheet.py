from PIL import Image, ImageDraw
from pathlib import Path
fs=sorted(Path('tmp/pdfs').glob('guide-*.png'))
thumbs=[]
for i,p in enumerate(fs):
    im=Image.open(p).convert('RGB')
    im.thumbnail((440,620))
    canvas=Image.new('RGB',(460,660),'white')
    canvas.paste(im,((460-im.width)//2,25))
    ImageDraw.Draw(canvas).text((10,5),f'Page {i+1}',fill='black')
    thumbs.append(canvas)
out=Image.new('RGB',(920,1980),'#dddddd')
for i,im in enumerate(thumbs):
    out.paste(im,((i%2)*460,(i//2)*660))
out.save('tmp/pdfs/contact.png')
