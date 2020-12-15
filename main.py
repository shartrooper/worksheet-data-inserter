#! python3
import sys
import getDatafromFiles as bt
import HGFExcelInserter as hgf
import os
from pathlib import Path
from pikepdf import _cpphelpers
from getmac import getmac

clients = ['EC:B1:D7:4F:1E:99', '6C:62:6D:91:1E:D2', '00:25:AB:AC:F9:E7', '40:A8:F0:A8:AF:81', '64:51:06:2D:BE:BC',
           '64:51:06:33:D3:C4', '00:25:AB:A4:19:73', '40:61:86:86:E3:EA', '40:61:86:7E:67:58', '64:51:06:3F:4B:FC',
           '64:51:06:31:ED:AA', '30:9C:23:0B:48:C4', 'F4:B5:20:1C:1B:C6']

getMac = logging.getLogger('getmac')
getMac.setLevel(logging.INFO)

# getmac.DEBUG = 0
client = getmac.get_mac_address()

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

if (len(sys.argv) >= 4 & (client.upper() in clients):
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