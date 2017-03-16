import os
import slate
import re
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ImageXL
import csv
import datetime
import barcode
from barcode.writer import ImageWriter
from PIL import Image

travelcust={}
travelcust["01 CTI (Cent)"]="01 CTI"
travelcust["02 Jabil Wolf"]="02 Wolfe"
travelcust["03 Assured MFG"]="04 TRU"
travelcust["04 TRU Machining"]="05 OTIS"
travelcust["05 OTIS (Ewa)"]="03 Assured"
travelcust["06 ThermoFisher (Ewa)"]="06 Thermo"
travelcust["07 Vecco (Luy Danh)"]="07 Vecco"
travelcust["08 Intevac (Luy Danh)"]="08 Intevac"
travelcust["09 Dung Marina"]="10-C"
travelcust["10 C.X.O (Sam)"]="09 Marina"
travelcust["11 Embraer (Sam)"]="11 Embraer"
travelcust["12 MW Technology - Cris (Sam)"]="12 MW Tech"

filename = "C:\PT&R\Customers\\10.Sam\C. X.O\SO\SO 31824\SO 31824.pdf"
txtfile = os.path.splitext(filename)[0]+'.txt'
out = open(txtfile,"w") 
with open(filename, 'rb') as f:
    doc = slate.PDF(f)
for page in doc[:2]:
    out.write(page)
out.close()
sodir = os.path.dirname(filename)
jobdir = os.path.dirname(sodir) # new
custdir = os.path.dirname(jobdir) # new
traveldir = sodir # new
custdir = os.path.dirname(sodir) # old, will be removed
traveldir = custdir + "\\Traveler" # old, will be removed
jobdir = custdir + "\\01 Job Documents"  # old, will be removed
faidir = custdir + "\\FAI"
shipdir = custdir + "\\Shipping"
custname = os.path.basename(custdir)
text_file = open(txtfile, "r")
lines = text_file.read().split('\n\n')
enum_lines = enumerate(lines)

found_item_no = 0
itemNameFound = 0
sosize=2
if (sosize > 2):
    adj1 = 0
    adj2 = 0
    lineSearch = 'Amount'
if (sosize == 2):
    adj1 = 7
    adj2 = 1
    lineSearch = 'Line'
if (sosize == 1):
    adj1 = 12
    adj2 = 1
    lineSearch = 'Item'


for iteration in enum_lines:
    index, item = iteration
    print (iteration)
    if  (item == "S.O. No."):
        sodate = lines[index+1]
        SO = lines[index+2]
        print sodate,SO
    if  (item == "P.O. No."):
        servicetype = lines[index+1]
        PO = lines[index+2]
        SOPO = "SO "+SO+" PO "+PO
        print servicetype, PO
    if  (item == "Ship Method"):
        paymentDue = lines[index+1]
        soShipDate = lines[index+2]
        print paymentDue, soShipDate
    if ( found_item_no > 0 ):
        if ( item.isdigit() ):
            itemNo = int(item)
            #print found_item_no,item
            if ( found_item_no == itemNo ):
                itemMax = itemNo
                found_item_no = itemNo + 1
            else:
                print itemMax
                itemNameFound = 1
                found_item_no = -1
        else:
            print itemMax
            itemNameFound = index
            found_item_no = -1
    if  (item == lineSearch ):
        #print item
        found_item_no = 1
    if ( itemNameFound ):
        index1 = index
        indexnext = index1 + itemMax
        itemNames = lines[index1:indexnext]
        print itemNames
        index1 = indexnext + adj1
        indexnext = index1 + itemMax
        itemDesc = lines[index1:indexnext]
        print itemDesc
        index1 = indexnext + adj2
        indexnext = index1 + itemMax
        itemQty = lines[index1:indexnext]
        print itemQty
        itemNameFound = 0
#
# Generate Directories Traveler
#
dirs = [jobdir,faidir,shipdir,traveldir]
for onedir in dirs:
    sopodir = onedir + "\\" + SOPO
    print sopodir
    if not os.path.exists(sopodir):
        os.makedirs(sopodir)
travelsodir = sopodir
print traveldir
for index,partNo in enumerate(itemNames):
    print index,partNo
    lineno = index + 1
    wb = load_workbook("c:\\data\\traveler.xlsx")
    ws = wb.active

    itemDescLines = itemDesc[index].split('\n')
    n = len(itemDescLines)
    if (n == 1):
        print "Description ",itemDescLines," not in correct format"
        exit()
    for i in range(n):
        if (i>0):
            desc = itemDescLines[i]
            if re.match(r'^Mat: ', desc):
                mtype = desc.replace("Mat: ", "", 1)
            if re.match(r'^Material: ', desc):
                mtype = desc.replace("Mat: ", "", 1)
            if re.match(r'^Fin: ', desc):
                fin = desc.replace("Fin: ", "", 1)
        else:       
            pdescription = itemDescLines[0]
    
    shipdate=soShipDate
    duedate=paymentDue
    quoteno="quoteno"
    #PO="PO 123"
    jobno=SO+"-LN"+str(lineno)
    drawno = "drawno"
    custno = custname
    qty_order = itemQty[index]
    #mtype = "M type"
    mlotno = "Lot 30304"
    outside = servicetype
    o_vendor = ""
    o_vendor_po = ""

    ws["C1"] = partNo 
    ws["C2"] = pdescription
    d = datetime.datetime.now()
    only_date, only_time = d.date(), d.time()
    ws['I1'].number_format = 'm/d/y'
    ws["I1"] = shipdate
    ws["I2"] = duedate
    ws["I3"] = quoteno
    ws["I4"] = PO
    ws["H7"] = outside
    ws["H8"] = o_vendor
    ws["H9"] = o_vendor_po
    ws["C5"] = jobno
    ws["C6"] = custno
    ws["C7"] = drawno
    ws["C8"] = only_date
    ws['C8'].number_format = 'm/d/y'
    ws["C10"] = qty_order
    ws["C13"] = mtype
    ws["C14"] = mlotno
    imgfile = travelsodir + "\\" + partNo
    pngfile = travelsodir + "\\" + partNo + ".png"
    CODE39 = barcode.get_barcode_class('code39')
    code39 = CODE39(partNo, writer=ImageWriter(),add_checksum=False)
    fullname = code39.save(imgfile)
    img = Image.open(fullname) # image extension *.png,*.jpg
    new_width  = 186
    new_height = 120
    img = img.resize((new_width, new_height), Image.ANTIALIAS)
    img.save(fullname)
    img1 = ImageXL(fullname)
    ws.add_image(img1,'L1')
    #img1 = Image('c:\\data\\logo1.png')
    #ws.add_image(img1,'L1')
    #img2 = Image('c:\\data\\logo2.png')
    #ws.add_image(img2,'L3')
    traveler = travelsodir + "\\T " + partNo + ".xlsx"
    wb.save(traveler)
    wb.close()  
