using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word= Microsoft.Office.Interop.Word;
using Visio= Microsoft.Office.Interop.Visio;
namespace OfficeHelper
{
    class Program
    {
        static void Main(string[] args)
        {
            try {
                ExcelHandler excelHandler = new ExcelHandler();
                excelHandler.processExcel();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
