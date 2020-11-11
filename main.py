#! python3
import sys
import electronbloodTestsInserterv2 as bt
import HGFExcelInserter as hgf


if len(sys.argv) >= 3:
    # Get arguments from command line.
    streamPath = sys.argv[1]
    glossaryPath = sys.argv[2]
    workSheetPath = ''
    if len(sys.argv) > 3:
        workSheetPath = sys.argv[3]
    glossary = open(glossaryPath, "r")
    dataStream = bt.GetDataStreamCollection(streamPath)
    gloss = bt.GetGlossary(glossary).getCollection()
    #print(gloss)
    collection = dataStream.getCollection()
    wb = bt.LoadOrCreateWorkBook(workSheetPath).getWorkBook()
    ws = wb.active
    currentWS = hgf.InsertDataInWorkSheet(ws, collection, "Resumen exámenes de laboratorio UPC HGF", gloss)
    currentWS.insertAndFormatHeaderData()
    currentWS.insertAndFormatDates()
    currentWS.insertTestResultData()
    currentWS.AddThinCellBorderStyle()
    header = currentWS.getHeaderFormat()
    if not workSheetPath:
        wb.save(header['Nombre'] + header['RUT'] + ".xlsx")
    else:
        wb.save(workSheetPath)
    bt.WriteLog('Worksheet updated/created!')
else:
    print(
        "Please, introduce at least existing filenames for test report and dictionary in command lines"
    )
    sys.exit(1)
