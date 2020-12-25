#! python3
import sys
import getDatafromFiles as bt
import HGFExcelInserter as hgf
import os
from pathlib import Path
from pikepdf import _cpphelpers

def newOrUpdatedWs():
    currentWS.insertAndFormatHeaderData()
    if currentWS.getErrorFlag():
        return True
    currentWS.insertAndFormatDates()
    if currentWS.getErrorFlag():
        return True
    if currentWS.isColCapReached():
        return 'capped!'
    currentWS.insertTestResultData()
    currentWS.addThinCellBorder()
    currentWS.setDataColumnDimensions()
    currentWS.removeColSurplus()
    hgf.AdjustCalciumValue(ws,currentWS.getCurrentDateCoordinates())
    hgf.CreateRecycleWorkSheet(wb,ws,4)

if len(sys.argv) >= 5:
    # Get arguments from command line.
    _,glossaryPath,workSheetPath,*streamPathCollection,savePath=sys.argv
    if workSheetPath == 'none':
        workSheetPath = ''
    print(os.path.isdir(savePath))
    if savePath == 'none' or not os.path.isdir(savePath):
        defaultPath = "\\blood tests data\\"
        if not os.path.isdir(os.getcwd() + defaultPath):
            Path("blood tests data").mkdir()
        savePath=os.getcwd() + defaultPath
    glossary = open(glossaryPath, "r")
    gloss = bt.GetGlossary(glossary).getCollection()
    wb = bt.LoadOrCreateWorkBook(workSheetPath).getWorkBook()
    ws = wb.active
    header=None
    human=''
    for streamPath in streamPathCollection:
        dataStream = bt.GetDataStreamCollection(streamPath)
        collection = dataStream.getCollection()
        if not collection:
            continue
        currentWS = hgf.InsertDataInWorkSheet(ws, collection, "Resumen exámenes de laboratorio UPC de Centros de Salud", gloss)
        catchError=newOrUpdatedWs()
        if catchError:
            continue
        if catchError == 'capped!':
            break
    header=currentWS.getHeaderFormat()
    if not workSheetPath:
        toSaveFilename=header['Nombre'] + header['RUT'] + ".xlsx"
        for root, dirs, files in os.walk(savePath):
            for i in range(1,len(files)+1):
                for filename in files:
                    if toSaveFilename == filename:
                        toSaveFilename= header['Nombre'] + header['RUT'] +"("+str(i)+")"+".xlsx"
        saveFile=savePath+toSaveFilename
        wb.save(saveFile)
        os.system(f'cmd /c "{saveFile}"')
    else:
        wb.save(workSheetPath)
        os.system(f'cmd /c "{workSheetPath}"')
    bt.WriteLog('Worksheet updated/created!')
else:
    print(
        "Please, introduce at least existing filenames for test reports and dictionary in command lines"
    )
    sys.exit(1)