#! python3
import sys
import electronbloodTestsInserterv2 as bt
import HGFExcelInserter as hgf
import os
from pathlib import Path
from pikepdf import _cpphelpers

def newOrUpdatedWs():
    currentWS = hgf.InsertDataInWorkSheet(ws, collection, "Resumen exámenes de laboratorio UPC de Centros de Salud", gloss)
    currentWS.insertAndFormatHeaderData()
    currentWS.insertAndFormatDates()
    currentWS.insertTestResultData()
    currentWS.addThinCellBorder()
    currentWS.setDataColumnDimensions()
    currentWS.removeColSurplus()
    hgf.AdjustCalciumValue(ws,currentWS.getCurrentDateCoordinates())
    hgf.CreateRecycleWorkSheet(wb,ws,4)
    header=currentWS.getHeaderFormat()
    return [header,os.getcwd() + '\\blood tests data\\' + header['Nombre'] + header['RUT'] + ".xlsx"]


if len(sys.argv) >= 3:
    # Get arguments from command line.
    streamPath = sys.argv[1]
    glossaryPath = sys.argv[2]
    workSheetPath = ''
    #glossaryPath,workSheetPath,*streamCollection=sys.argv
    if len(sys.argv) > 3:
        workSheetPath = sys.argv[3]
    glossary = open(glossaryPath, "r")
    dataStream = bt.GetDataStreamCollection(streamPath)
    gloss = bt.GetGlossary(glossary).getCollection()
    collection = dataStream.getCollection()
    wb = bt.LoadOrCreateWorkBook(workSheetPath).getWorkBook()
    ws = wb.active
    header,human=newOrUpdatedWs()
    if not workSheetPath:
        if not os.path.isdir(os.getcwd() + '\\blood tests data'):
            Path("blood tests data").mkdir()
        wb.save(os.getcwd() + '\\blood tests data\\' + header['Nombre'] + header['RUT'] + ".xlsx")
        os.system(f'cmd /c "{human}"')
    else:
        wb.save(workSheetPath)
        os.system(f'cmd /c "{workSheetPath}"')
    bt.WriteLog('Worksheet updated/created!')
else:
    print(
        "Please, introduce at least existing filenames for test report and dictionary in command lines"
    )
    sys.exit(1)
