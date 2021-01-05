#! python3
import sys
import getDatafromFiles as bt
import HGFExcelInserter as hgf
import os
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
    dataStream = bt.GetDataStreamCollection(streamPathCollection[0])
    collection = dataStream.getCollection()
    if not collection:
        raise Exception("An exception happened, check log file for details.")
    saveFileRoutes=hgf.SaveFileRouteExplorer(savePath,workSheetPath,collection)
    workSheetPath =saveFileRoutes.getWSPath()
    saveFileName = saveFileRoutes.getSaveFileName()
    savePath = saveFileRoutes.getSaveRoute()
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
        saveFile=savePath+saveFileName
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