# bloodTestsInserter.pyv2 get keywords from test reports with a specific format and inserts results in a .xlsx
"""
-   takes either a .txt or .pdf, a glossary of keywords and output filenames from command line
-   e.g. bloodTestsInserter.py <test report filename> <glossary> <output worksheet>
-   If a worksheet isn't loaded, creates a new one from the glossary.
-   Only inserts test results that exists in glossary
"""
import sys, PyPDF2, logging, re, datetime
from openpyxl import Workbook, load_workbook

logging.basicConfig(
    level=logging.DEBUG, format="%(asctime)s -  %(levelname)s -  %(message)s"
)


def openFiles(filenames):
    try:
        fileExtRegex = re.compile(r"txt|pdf")
        found = fileExtRegex.search(filenames[1])
        txtStream = []
        if found:
            if found.group(0) == "txt":
                txtFile = open(filenames[1])
                content = txtFile.read()
                txtStream.append(content)
            elif found.group(0) == "pdf":
                pdfFileObj = open(filenames[1], "rb")
                pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
                for numPage in range(0, pdfReader.getNumPages()):
                    pageObj = pdfReader.getPage(numPage)
                    txtStream.append(pageObj.extractText())
                    # logging.debug(txtStream[numPage])
                pdfFileObj.close()
        else:
            raise Exception("Not a valid file ext")

        dictFile = open(filenames[2], "r")
        dictStrings = dictFile.readlines()
        dictio = {}
        category = ""
        for i in range(0, len(dictStrings)):
            if dictStrings[i].isupper():
                category = dictStrings[i].strip()
                dictio[category] = []
                continue
            dictio[category].append(dictStrings[i].strip().lower())
            # logging.debug(dictio[category])
        if len(filenames) > 3:
            wb = load_workbook(filenames[3])
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "new sheet"
        dictFile.close()
    except Exception as err:
        print("An exception happened: " + str(err))
        sys.exit(1)
    return [txtStream, dictio, wb]


def strIntoUsefulRegex(testname):
    charList = ["\.", "\(", "\)"]
    for char in charList:
        charRegex = re.compile(rf"{char}")
        testname = charRegex.sub(char, testname)
    return testname


def getTotalLength(glossary):
    count = 0
    for category in glossary:
        count += len(glossary[category])
    return count


if len(sys.argv) >= 3:
    # Get arguments from command line.
    stream, dictionary, wb = openFiles(sys.argv)
else:
    print(
        "Please, introduce at least existing filenames for test report and dictionary in command lines"
    )
    sys.exit(1)

# If apply, create workBook with testnames rows
dateRegex = re.compile(
    r"Fecha Recepci[oó]n\s?:?\s?(([0-3]?[1-9])[\/-]([0-1]?[1-9])[\/-]([1-2][0-9]{3}))\s(\d{2}:\d{2})",
    re.IGNORECASE,
)

for page in stream:
    if dateRegex.search(page):
        reportDate = dateRegex.search(page)
        break
    else:
        raise Exception("Date not found")

wsDate = datetime.date(
    int(reportDate.group(4)), int(reportDate.group(3)), int(reportDate.group(2))
).strftime("%d/%m/%y")
wsTime = reportDate.group(5)
ws = wb.active
dateCoordinates = ""


if ws.title == "new sheet":
    ws["A1"].value = "Fecha"
    ws["A2"].value = "Hora"
    rowPos = 3
    for category in dictionary:
        for j, testname in enumerate(dictionary[category], start=rowPos):
            currentCell = ws.cell(row=j, column=1, value=testname)
            # logging.debug(str(currentCell.row)+' with value '+currentCell.value)
            rowPos += 1
    ws.title = "RESULTADOS"
else:
    # Add new testname as a WorkSheet row
    rowPos = 3
    dictLen = getTotalLength(dictionary) + 2
    if dictLen > ws.max_row:
        for category in dictionary:
            for j, testname in enumerate(dictionary[category], start=rowPos):
                if ws.cell(row=j, column=1).value != testname:
                    ws.insert_rows(j)
                    ws.cell(row=j, column=1, value=testname)
                rowPos += 1
# Add date value to a column

for col in ws.iter_cols(min_col=2, max_col=999999):
    if dateCoordinates:
        break
    for cell in col:
        timeCell = cell.coordinate[0] + str(2)
        if not cell.value:
            cell.value = wsDate
            ws[timeCell].value = wsTime
            dateCoordinates = cell.coordinate
        elif cell.value == wsDate and ws[timeCell].value == wsTime:
            dateCoordinates = cell.coordinate
        elif cell.value != wsDate or ws[timeCell].value != wsTime:
            wsDateParams = wsDate.split("/")
            colDateParams = cell.value.split("/")
            colTimeCell = ws[timeCell].value
            d1 = datetime.datetime(
                int(wsDateParams[2]),
                int(wsDateParams[1]),
                int(wsDateParams[0]), int(wsTime[:2]), int(wsTime[3:5])
            )
            d2 = datetime.datetime(
                int(colDateParams[2]),
                int(colDateParams[1]),
                int(colDateParams[0]),
                int(colTimeCell[:2]),
                int(colTimeCell[3:5])
            )
            if d1 < d2:
                ws.insert_cols(cell.column, amount=1)
                ws.cell(row=cell.row, column=cell.column - 1, value=wsDate)
                dateCoordinates = ws.cell(
                    row=cell.row, column=cell.column - 1
                ).coordinate
                newTimeCoordinate=dateCoordinates[0]+str(2)
                ws[newTimeCoordinate].value=wsTime
        break

# Match and insert existing testname results in textStream

for category in dictionary:
    for testname in dictionary[category]:
        for row in ws.rows:
            for cell in row:
                if testname == cell.value:
                    for page in stream:
                        reTestname = strIntoUsefulRegex(testname)
                        testRegex = re.compile(
                            rf"{reTestname}[\s]*[:]?[\s]*([-*]?\d+\.?\d*)",
                            re.IGNORECASE,
                        )
                        if testRegex.search(page):
                            testResult = testRegex.search(page)
                            ws[
                                dateCoordinates[0] + str(cell.row)
                            ].value = testResult.group(1)
                            break
                    break

# Save WorkBook on a file

wb.save("report" + ".xlsx")
print("Workbook updated or saved in local dir!")
