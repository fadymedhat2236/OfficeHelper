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
            Word.Application wordApp= new Word.Application();
            try {
                WordHandler wordHandler = new WordHandler();
                VisioHandler visioHandler = new VisioHandler();
                Word.Document doc;
                //for thr SCI
                /*
                doc = wordHandler.generateSCI(wordApp);
                doc.Save();
                //copy small chunk to avoid large clipboard objects warning message on close
                doc.Sections[1].Range.Copy();
                doc.Close();*/

                //for the SPI
                /*doc = wordHandler.generateSPI(wordApp);
                doc.Save();
                //copy small chunk to avoid large clipboard objects warning message on close
                doc.Sections[1].Range.Copy();
                doc.Close();*/

                //for the DTD
                doc = wordHandler.generateDTD(wordApp);
                doc.Save();
                visioHandler.addSequanceDiagram(ref doc);
                visioHandler.addRequestFlowDiagram(ref doc);
                visioHandler.addResponseFlowDiagram(ref doc);
                //copy small chunk to avoid large clipboard objects warning message on close
                doc.Sections[1].Range.Copy();
                doc.Close();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                wordApp.Quit();
            }

        }
    }
}
