from PIL import Image
from pdf2image import convert_from_path
import os, glob, pytesseract, win32com.client, tempfile, datetime, re
import PyPDF2 as pdf2

mos = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

def isCheque(arr):
    return any('Cheque' in s for s in arr)

def parseDate(inp):
    dates = list(filter(re.compile('[a-zA-Z\s|.]*[0-9]*/[0-9]*/[0-9]*').match, inp))
    for i in range(len(dates)):
        dates[i] = re.findall('[0-9]*/[0-9]*/[0-9]*', dates[i])[0]
    date = max(dates).replace('/', '')

    month = int(date[2:4])
    day = date[0:2]
    year = date[4:]
    return mos[month - 1] + day + year

def getFolder(title):
    date = title.split('_')[2]
    month = mos.index(date[:3]) + 1
    year = date[5:]
    if (int(month) > 9):
        year = str(int(year) + 1)
    return "HSC AP Oct 1 to Sept 30 " + year + "\\" + title.split("_")[0]

def makeTitle(arr):
    indices = re.search('[a-zA-Z0-9_&]+( [a-zA-Z0-9_&\/-]+)*', arr[1])
    name = arr[1][indices.start():indices.end()].replace('/', '-')
    date = parseDate(arr)

    potentialCheques = list(filter(re.compile('[0-9a-zA-Z\s/\//|.]*01[0-9]{4,}').match, arr))
    chequeNo = '#' + re.findall('01[0-9]{4,}', potentialCheques[0])[0][1:]

    pAmt = list(filter(re.compile('[0-9a-zA-Z\s/\//|.]*([0-9]+\,)*[0-9]+\.[0-9]+').match, arr))
    for i in range(len(pAmt)):
        pAmt[i] = re.search('([0-9]+\,)*[0-9]+\.[0-9]+', pAmt[i])[0]
    amt = max(pAmt)

    title = name + '_' + chequeNo + '_' + date + '_$' + amt + '.pdf'
    return title

def saveTempFile(fil, tempdir):
    a.SaveAsFile(tempdir + "\\" + str(a))
    print('saved temporary file to ' + tempdir)
    return tempdir + "\\" + str(a)

def getFirstPage(fil, tempdir):
    print('getting first page of pdf')
    with open(fil, 'rb') as f:
        pdf = pdf2.PdfFileReader(f)
        output = pdf2.PdfFileWriter()
        output.addPage(pdf.getPage(0))

        writePath = os.path.join(tempdir + "\\tempFirstPage.pdf")
        with open(writePath, 'wb') as outputStream:
            output.write(outputStream)
            return tempdir + "\\tempFirstPage.pdf"

def savePDF(original, firstPage):
    print('converting to image')
    doc = convert_from_path(firstPage, 500)
    text = pytesseract.image_to_string(doc[0]).split('\n')
    text = list(filter(None, text))
    print(text)

    if (not isCheque(text)):
        return False

    print('making title')
    title = makeTitle(text)

    print('getting folder')
    saveDir = getFolder(title)
    if not os.path.exists("Z:\\HSC Holdings AP Files\\" + saveDir):
        print('creating folder')
        os.makedirs("Z:\\HSC Holdings AP Files\\" + saveDir)
    
    with open(original, 'rb') as f:
        pdf = pdf2.PdfFileReader(f)
        output = pdf2.PdfFileWriter()
        output.appendPagesFromReader(pdf)
        writePath = os.path.join("Z:\\HSC Holdings AP Files\\" + saveDir, title)

        print('saving pdf')
        with open(writePath, 'wb') as outputStream:
            output.write(outputStream)
        print('saved to ' + writePath)
        return True

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
other = outlook.Folders.Item(1).Folders["Other"]
done = outlook.Folders.Item(1).Folders["Saved"]

while (True):
    inbox = outlook.GetDefaultFolder(6)
    for m in inbox.Items:
        if (m.UnRead == True and m.Subject == 'Attached Image'):
            for a in m.Attachments:
                with tempfile.TemporaryDirectory() as tempdir:
                    filePath = saveTempFile(a, tempdir)
                    firstPage = getFirstPage(filePath, tempdir)
                    if (savePDF(filePath, firstPage)):
                        m.UnRead = False
                        m.Move(done)
                    else:
                        print("not a cheque")
                        m.Move(other)
                    continue