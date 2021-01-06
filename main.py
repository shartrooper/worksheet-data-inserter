#! python3
import logging
import sys
import getDatafromFiles as bt
import HGFExcelInserter as hgf
import os
from pikepdf import _cpphelpers
from getmac import getmac

clients = ['EC:B1:D7:4F:1E:99', '6C:62:6D:91:1E:D2', '00:25:AB:AC:F9:E7', '40:A8:F0:A8:AF:81', '64:51:06:2D:BE:BC',
           '64:51:06:33:D3:C4', '00:25:AB:A4:19:73', '40:61:86:86:E3:EA', '40:61:86:7E:67:58', '64:51:06:3F:4B:FC',
           '64:51:06:31:ED:AA', '30:9C:23:0B:48:C4', 'F4:B5:20:1C:1B:C6']

getMac = logging.getLogger('getmac')
getMac.setLevel(logging.INFO)

# getmac.DEBUG = 0
client = getmac.get_mac_address()

def newOrUpdatedWs(sheet):
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
    hgf.AdjustCalciumValue(sheet,currentWS.getCurrentDateCoordinates())

if (len(sys.argv) >= 5) and (client.upper() in clients):
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
    human=''
    recycle=hgf.CreateRecycleWorkSheet(wb,ws)
    recycle.removeRecycleSheet()
    for streamPath in streamPathCollection:
        dataStream = bt.GetDataStreamCollection(streamPath)
        collection = dataStream.getCollection()
        if not collection:
            continue
        for sheet in wb.worksheets:
            currentWS = hgf.InsertDataInWorkSheet(sheet, collection, "Resumen exámenes de laboratorio UPC de Centros de Salud", gloss)
            catchError=newOrUpdatedWs(sheet)
            if catchError and catchError != "capped!":
                break
            if catchError == 'capped!':
                wsIndex=wb.index(sheet)
                if wsIndex == len(wb.worksheets)-1:
                    nws=wb.copy_worksheet(sheet)
                    for col in nws.iter_cols(min_col=2, min_row=4, max_col=12):
                        for cell in col:
                            cell.value=''
                    currentWS = hgf.InsertDataInWorkSheet(nws, collection, "Resumen exámenes de laboratorio UPC de Centros de Salud", gloss)
                    newOrUpdatedWs(nws)
                    recycle=hgf.CreateRecycleWorkSheet(wb,nws)
                continue
            recycle=hgf.CreateRecycleWorkSheet(wb,sheet)
            break
    recycle.createRecycleSheet()
    if not workSheetPath:
        saveFile=savePath+saveFileName
        wb.save(saveFile)
        os.system(f'cmd /c "{saveFile}"')
    else:
        wb.save(workSheetPath)
        os.system(f'cmd /c "{workSheetPath}"')
    bt.WriteLog('Worksheet updated/created!')
else:
    if not client.upper() in clients:
        print("UNAUTHORIZED USER!")
        sys.exit(1)
    print(
        "Please, introduce at least existing filenames for test reports and dictionary in command lines"
    )