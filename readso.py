import os
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import csv
import datetime

filename = "C:\PT&R\Customers\\10.Sam\C. X.O\SO\SO 31839.txt"
sodir = os.path.dirname(filename)
custdir = os.path.dirname(sodir)
traveldir = custdir + "\\Traveler"
jobdir = custdir + "\\Job Documents"
faidir = custdir + "\\FAI"
shipdir = custdir + "\\Shipping"
custname = os.path.basename(custdir)
text_file = open(filename, "r")
lines = text_file.read().split('\n\n')
enum_lines = enumerate(lines)

for iteration in enum_lines:
    index, item = iteration
    if  (item == "S.O. No."):
        index = index + 1
        break
sodate = lines[index:index+1]
SO = lines[index+1:index+2]
print sodate,SO

for iteration in enum_lines:
    index, item = iteration
    if  (item == "P.O. No."):
        index = index + 1
        break
servicetype = lines[index:index+1]
PO = lines[index+1:index+2]
SOPO = "SO "+SO[0]+" PO "+PO[0]
print servicetype, PO

for iteration in enum_lines:
    index, item = iteration
    if  (item == "Ship Method"):
        index = index + 1
        break
paymentDue = lines[index:index+1]
soShipDate = lines[index+1:index+2]
print paymentDue, soShipDate

found_item_no = 0
itemNameFound = 0
for iteration in enum_lines:
    #print (iteration)
    index, item = iteration
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
    if  (item == "Amount"):
        found_item_no = 1
    if ( itemNameFound ):
        break
indexnext = index + itemMax
itemNames = lines[index:indexnext]
print itemNames
index = indexnext
indexnext = index + itemMax
itemDesc = lines[index:indexnext]
print itemDesc
index = indexnext
indexnext = index + itemMax
itemQty = lines[index:indexnext]
print itemQty
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
    pdescription = itemDescLines[0]
    mtype = itemDescLines[1]
    
    shipdate=soShipDate[0]
    duedate=paymentDue[0]
    quoteno="quoteno"
    #PO="PO 123"
    jobno=SO[0]+"LN"+str(lineno)
    drawno = "drawno"
    custno = custname
    qty_order = itemQty[index]
    #mtype = "M type"
    mlotno = "Lot 30304"
    outside = servicetype[0]
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
    ws["I4"] = PO[0]
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
    #img1 = Image('c:\\data\\logo1.png')
    #ws.add_image(img1,'L1')
    #img2 = Image('c:\\data\\logo2.png')
    #ws.add_image(img2,'L3')
    traveler = travelsodir + "\\T " + partNo + ".xlsx"
    wb.save(traveler)
    wb.close()       
