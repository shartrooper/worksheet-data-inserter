import unittest
import electronbloodTestsInserterv2 as testBT
import HGFExcelInserter as testHGF
import os.path
from os import path


testStr="""Hematología*\n
Hb: Hemoglobina\n
Hct: Hematocrito\n
VCM: V.C.M. (Volumen Medio)\n
HCM: H.C.M. (Hemoglobina Media)\n
CHCM: C.H C M (Concentracion Media)\n
A.D.E.: A.D.E.\n
Leucocitos: Recuento Leucocitos\n
RAN totales: Neutrofilos totales\n
RAL totales: Linfocitos totales\n
MN totales: Monocitos totales\n
EOS totales: Eosinofilos totales\n
Plaquetas: Plaquetas (recuento)\n
VHS: vhs\n
ELECTROLITOS PLASMÁTICOS*\n
Cloro: cloro\n
Potasio: potasio\n
Sodio: sodio\n
Ca Iónico: calcio Iónico\n
PERFIL BIOQUÍMICO*\n
Ferritina: ferritinina\n
Ácido Urico: acido urico\n
Calcio total: calcio\n
LDH: lactato deshidrogenasa\n
Fostoro: fosforo\n
BUN:Nitrogeno ureico\n
Urea: urea\n
Albumina: albumina\n
Proteinas Totales: proteinas totales\n
PCR: Proteinas c reactiva\n
Creatinina: Creatinina\n
PRUEBAS HEPÁTICAS*\n
Bil Total: bilirrubina total\n
Bil Directa: bilirrubina directa\n
FA: fosfatasa alcalina\n
GGT: ggt\n
GOT: got\n
GPT: gpt\n
CINÉTICA DEL FIERRO*\n
Fierro: fierro\n
U.I.B.C: u.i.b.c\n
T.I.B.C: t.i.b.c\n
% Sat Transferrina: %  saturación de transferrina\n
transferrina: transferrina\n
VOLUMEN FILTRADO GLOMERULAR*\n
TP: porcentaje tp\n
INR: inr\n
TTPA: ttpa\n
GASES EN SANGRE*\n
pH: ph\n
PCO2: p co2\n
PO2: p o2\n
HCO3: hco3\n
Ex. Base: ex. Base\n
SatPO2: saturacion de o2\n
"""

class testCollection(unittest.TestCase):
    """
        Testing for General Collection Creator classes
    """

    def test_reportCollectionOutput(self):
        outputCollection = testBT.GetDataStreamCollection("25-07.pdf")
        self.assertEqual(
            type(outputCollection.getCollection()) is list,
            True,
            "Collection has to be a list",
        )

    def test_keyWordGlossaryCollectionOutput(self):
        outputCollection = testBT.GetGlossary(testStr)
        self.assertEqual(
            type(outputCollection.getCollection()) is dict,
            True,
            "Collection has to be a dictionary"
        )

class testHGFClasses(unittest.TestCase):
    """
        Testing for HGF Hospital format classes
    """
    def test_getHeaderContent(self):
        dataStream = testBT.GetDataStreamCollection("25-07.pdf")
        collection= dataStream.getCollection()
        testHeader= testHGF.GetHeaderContent(collection,"This is a test header title")
        self.assertEqual(type(testHeader.getHeaderFormat()) is dict,True, "returned Collection has to be a dictionary")
        self.assertEqual(type(testHeader.getRowPosition()) is int,True,"length value is int")

    def test_HeaderDataInsert(self):
        dataStream= testBT.GetDataStreamCollection("25-07.pdf")
        gloss= testBT.GetGlossary(testStr).getCollection()
        collection= dataStream.getCollection()
        wb=testBT.LoadOrCreateWorkBook('').getWorkBook()
        ws=wb.active
        testwsHeader= testHGF.InsertDataInWorkSheet(ws,collection,"This is a test header title",gloss)
        testwsHeader.insertAndFormatHeaderData()
        wb.save("report" + ".xlsx")
        self.assertEqual(path.exists('report.xlsx'),True,"Report file is created")
    
    def test_HeaderAndDateTimeDataInsert(self):
        dataStream= testBT.GetDataStreamCollection("25-07.pdf")
        gloss= testBT.GetGlossary(testStr).getCollection()
        collection= dataStream.getCollection()
        wb=testBT.LoadOrCreateWorkBook('').getWorkBook()
        ws=wb.active
        testWS= testHGF.InsertDataInWorkSheet(ws,collection,"This is a test header title",gloss)
        testWS.insertAndFormatHeaderData()
        testWS.insertAndFormatDates()
        wb.save("report2" + ".xlsx")
        self.assertEqual(path.exists('report2.xlsx'),True,"Report file is created")
    
    def test_LoadingANewWorkSheetDate(self):
        dataStream= testBT.GetDataStreamCollection("25-07.pdf")
        gloss= testBT.GetGlossary(testStr).getCollection()
        collection= dataStream.getCollection()
        wb=testBT.LoadOrCreateWorkBook('').getWorkBook()
        ws=wb.active
        testWS= testHGF.InsertDataInWorkSheet(ws,collection,"This is a test header title",gloss)
        testWS.insertAndFormatHeaderData()
        testWS.insertAndFormatDates()
        wb.save("report3" + ".xlsx")
        #Insert another PDF data in already created WorkSheet
        dataStream= testBT.GetDataStreamCollection("26-07.pdf")
        collection= dataStream.getCollection()
        wb=testBT.LoadOrCreateWorkBook('report3.xlsx').getWorkBook()
        ws=wb.active
        testWS= testHGF.InsertDataInWorkSheet(ws,collection,"",gloss)
        testWS.insertAndFormatHeaderData()
        testWS.insertAndFormatDates()
        wb.save("report3" + ".xlsx")
        self.assertEqual(path.exists('report3.xlsx'),True,"Report file is created")
    
    def test_CompleteDataInsertion(self):
        dataStream= testBT.GetDataStreamCollection("29-07.pdf")
        gloss= testBT.GetGlossary(testStr).getCollection()
        collection= dataStream.getCollection()
        wb=testBT.LoadOrCreateWorkBook('').getWorkBook()
        ws=wb.active
        testWS= testHGF.InsertDataInWorkSheet(ws,collection,"This is a test header title",gloss)
        testWS.insertAndFormatHeaderData()
        testWS.insertAndFormatDates()
        testWS.insertTestResultData()
        testWS.AddThinCellBorderStyle()
        wb.save("report4" + ".xlsx")
        self.assertEqual(path.exists('report4.xlsx'),True,"Report file is created")

if __name__ == "__main__":
    unittest.main()
