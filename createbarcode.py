import barcode
from barcode.writer import ImageWriter
CODE39 = barcode.get_barcode_class('code39')
code39 = CODE39('CHAUN', writer=ImageWriter(),add_checksum=False)
pngfile='C:\data\chaup'
filename = code39.save(pngfile)
filename