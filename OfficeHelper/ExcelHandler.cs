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
        Dictionary<string, string> constants;
        public ExcelHandler()
        {
            excelApp = new Excel.Application();
            wordHandler = new WordHandler();
            visioHandler = new VisioHandler();
            constants = new Dictionary<string, string>();
            getConstants();
        }
        private int getRowsCount(Excel.Worksheet ws)
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
            for (int i = 1; i <= 14; i++)
            {
                string key = (string)(xlRange.Cells[1, i] as Excel.Range).Value2.ToString();
                string value = (string)(xlRange.Cells[rowNumber, i] as Excel.Range).Value2.ToString();
                dict.Add(key,value);
            }
            makeSCI = (string)(xlRange.Cells[rowNumber, 15] as Excel.Range).Value2.ToString();
            makeSPI = (string)(xlRange.Cells[rowNumber, 16] as Excel.Range).Value2.ToString();
            makeDTD = (string)(xlRange.Cells[rowNumber, 17] as Excel.Range).Value2.ToString();
            return dict;
        }
        private List<int> getTablesNumbers(string tablesNumbers)
        {
            List<int> l = new List<int>();
            char[] spearator = { ',' };
            String[] strlist = tablesNumbers.Split(spearator);
            foreach (string y in strlist)
            {
                l.Add(Int32.Parse(y));
                //Console.WriteLine(Int32.Parse(y));
            }
            return l;
        }
        private void getConstants()
        {
            Excel.Workbook wb = excelApp.Workbooks.Open(Directory.GetCurrentDirectory() + "\\Constants.xlsx");
            Excel.Worksheet ws = wb.Worksheets[3];
            int noRows = getRowsCount(ws);
            Excel.Range xlRange = ws.UsedRange;
            for (int i = 1; i <=noRows; i++)
            {
                string key = (string)(xlRange.Cells[i, 1] as Excel.Range).Value2.ToString();
                string value = (string)(xlRange.Cells[i, 2] as Excel.Range).Value2.ToString();
                constants.Add(key,value);
            }
        }
        public void processExcel()
        {
            Excel.Workbook wb = excelApp.Workbooks.Open(Directory.GetCurrentDirectory() + "\\Constants.xlsx");
            Excel.Worksheet ws = wb.Worksheets[1];
            int rowsNumber = getRowsCount(ws);
            for(int i = 0; i < rowsNumber-1; i++)
            {
                string makeSCI,makeSPI, makeDTD;
                makeSCI = makeSPI = makeDTD = "";
                Dictionary<string, string> dict = getConstants(ws,i + 2,ref makeSCI,ref makeSPI,ref makeDTD);
                if (makeSCI=="1")
                {
                    Word.Document doc= wordHandler.generateDocument(ref dict,ref constants,"SCI");
                    doc.Save();
                    doc.Close();
                }
                if (makeSPI=="1")
                {
                    Word.Document doc= wordHandler.generateDocument(ref dict,ref constants,"SPI");
                    doc.Save();
                    doc.Close();
                }
                if (makeDTD=="1")
                {
                    Word.Document doc=wordHandler.generateDocument(ref dict,ref constants,"DTD");
                    visioHandler.addSequanceDiagram(ref doc);
                    visioHandler.addRequestFlowDiagram(ref doc);
                    visioHandler.addResponseFlowDiagram(ref doc);

                    //copying old table data
                    //Request 9 Response 10
                    List<int> OldRequestTablesNumbers = getTablesNumbers(dict["OldRequestTablesNumbers"]);
                    List<int> OldResponseTablesNumbers = getTablesNumbers(dict["OldResponseTablesNumbers"]);
                    wordHandler.copyTableData(OldRequestTablesNumbers, 10, ref dict,ref constants,ref doc);
                    wordHandler.copyTableData(OldResponseTablesNumbers, 14, ref dict,ref constants,ref doc);
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
