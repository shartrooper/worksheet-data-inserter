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

if len(sys.argv) >= 4:
    # Get arguments from command line.
    _,glossaryPath,workSheetPath,*streamPathCollection=sys.argv
    if workSheetPath == 'none':
        workSheetPath = ''
    # TODO: don't overwrite files.
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
    human=os.getcwd() + '\\blood tests data\\' + header['Nombre'] + header['RUT'] + ".xlsx"
    if not workSheetPath:
        if not os.path.isdir(os.getcwd() + '\\blood tests data'):
            Path("blood tests data").mkdir()
        toSaveFilename=header['Nombre'] + header['RUT'] + ".xlsx"
        for root, dirs, files in os.walk(os.getcwd()+'\\blood tests data'):
            for i in range(1,len(files)+1):
                for filename in files:
                    if toSaveFilename == filename:
                        toSaveFilename= header['Nombre'] + header['RUT'] +"("+str(i)+")"+".xlsx"
        saveFile=os.getcwd() + '\\blood tests data\\' + toSaveFilename
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