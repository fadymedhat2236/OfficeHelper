using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
namespace OfficeHelper
{
    class ExcelHandler
    {
        Excel.Application excelApp;
        WordHandler wordHandler;
        VisioHandler visioHandler;
        public ExcelHandler()
        {
            excelApp = new Excel.Application();
            wordHandler = new WordHandler();
            visioHandler = new VisioHandler();
        }
        private int getRowsNumber(Excel.Worksheet ws)
        {
            try
            {
                Excel.Range lastCell = ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                return lastCell.Row;
            }
            catch(Exception exp)
            {
                throw exp;
            }
        }
        private Dictionary<string,string> getConstants(Excel.Worksheet ws, int rowNumber,ref string makeSCI, ref string makeSPI, ref string makeDTD)
        {
            Excel.Range xlRange = ws.UsedRange;
            Dictionary<string, string> dict = new Dictionary<string, string>();
            for (int i = 1; i <= 11; i++)
            {
                string key = (string)(xlRange.Cells[1, i] as Excel.Range).Value2.ToString();
                string value = (string)(xlRange.Cells[rowNumber, i] as Excel.Range).Value2.ToString();
                dict.Add(key,value);
            }
            makeSCI = (string)(xlRange.Cells[rowNumber, 12] as Excel.Range).Value2.ToString();
            makeSPI = (string)(xlRange.Cells[rowNumber, 13] as Excel.Range).Value2.ToString();
            makeDTD = (string)(xlRange.Cells[rowNumber, 14] as Excel.Range).Value2.ToString();
            return dict;
        }
        public void processExcel()
        {
            Excel.Workbook wb = excelApp.Workbooks.Open(Directory.GetCurrentDirectory() + "\\Constants.xlsx");
            Excel.Worksheet ws = wb.Worksheets[1];
            int rowsNumber = getRowsNumber(ws);
            for(int i = 0; i < rowsNumber-1; i++)
            {
                string makeSCI,makeSPI, makeDTD;
                makeSCI = makeSPI = makeDTD = "";
                Dictionary<string, string> dict = getConstants(ws,i + 2,ref makeSCI,ref makeSPI,ref makeDTD);
                if (makeSCI=="1")
                {
                    Word.Document doc= wordHandler.generateDocument(ref dict,"SCI");
                    doc.Save();
                    doc.Close();
                }
                if (makeSPI=="1")
                {
                    Word.Document doc= wordHandler.generateDocument(ref dict,"SPI");
                    doc.Save();
                    doc.Close();
                }
                if (makeDTD=="1")
                {
                    Word.Document doc=wordHandler.generateDocument(ref dict,"DTD");
                    visioHandler.addSequanceDiagram(ref doc);
                    visioHandler.addRequestFlowDiagram(ref doc);
                    visioHandler.addResponseFlowDiagram(ref doc);
                    doc.Save();
                    //copy small chunk to avoid large clipboard objects warning message on close
                    doc.Sections[1].Range.Copy();
                    doc.Close();
                }
            }
        }
        ~ExcelHandler()
        {
            excelApp.Quit();
        }
    }
}
