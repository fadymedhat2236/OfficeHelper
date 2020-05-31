using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace OfficeHelper
{
    class WordHandler
    {
        Word.Application wordApp = new Word.Application();
        
        public Word.Document generateDocument(ref Dictionary<string, string> dict,ref Dictionary<string,string> constantsDict,string documentName)
        {
            try
            {
                // Copy the template
                string DocFileName = constantsDict["Saving Path"] + "\\"+documentName+"\\" + "INT-002-004-" + dict["ServiceID"] + "-"+documentName+" " + dict["Subject"] + ".docx";
                File.Copy(Directory.GetCurrentDirectory() + "\\Templates\\" + dict["ServiceFlow"] + "\\Template_"+documentName+".docx", DocFileName, true);
                Word.Document doc = wordApp.Documents.Open(DocFileName);

                //setting the document properties
                doc.BuiltInDocumentProperties[Word.WdBuiltInProperty.wdPropertySubject].Value = dict["Subject"];
                doc.CustomDocumentProperties["Publish Date"].value = dict["PublishDate"];
                doc.CustomDocumentProperties["Review Date"].value = dict["ReviewDate"];
                doc.CustomDocumentProperties["Service Canonical Name"].value = dict["ServiceCanonicalName"];
                doc.CustomDocumentProperties["Service SubCategory"].value = dict["ServiceSubCategory"];
                doc.CustomDocumentProperties["Service ID"].value = dict["ServiceID"];
                if (documentName == "DTD")
                {
                    doc.CustomDocumentProperties["SSL Client Crypto Profile"].value = dict["SSLClientCryptoProfile"];
                    doc.CustomDocumentProperties["XML Manager"].value = dict["XMLManager"];
                    doc.CustomDocumentProperties["Backend Name"].value = dict["BackendName"];
                }
               
                //updating the documen
                updateDocument(ref doc);
                Console.WriteLine(DocFileName + " generated");
                return doc;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void addSCIErrorMapping(ref List<List<string>> errorMappng, ref Word.Document doc,string serviceFlow)
        {
            int tableNumber = 6;
            Word.Table ErrorTable = doc.Tables[tableNumber];
            int index=0;
            if(serviceFlow == "Inbound")
                index = 24;
            else if(serviceFlow == "Outbound")
                index = 27;
            foreach (List<string> errorRow in errorMappng)
            {
                if (serviceFlow == "Inbound")
                {
                    ErrorTable.Rows.Add(ErrorTable.Rows[index]);
                    //Status Code
                    Word.Cell cell = ErrorTable.Cell(index, 1);
                    cell.Range.Text = errorRow[1];
                    //Description
                    cell = ErrorTable.Cell(index, 2);
                    cell.Range.Text = errorRow[3];
                }
                else if (serviceFlow == "Outbound")
                {
                    ErrorTable.Rows.Add(ErrorTable.Rows[1]);
                    //Return Code
                    Word.Cell cell = ErrorTable.Cell(index, 1);
                    cell.Range.Text = errorRow[1];
                    //Reason Code
                    cell = ErrorTable.Cell(index, 2);
                    cell.Range.Text = errorRow[2];
                    //Description
                    cell = ErrorTable.Cell(index, 3);
                    cell.Range.Text = errorRow[3];
                }
                index++;
            }
            ErrorTable.Rows[index].Delete();
            Console.WriteLine("SCI Error Table Added");
        }
        public void addSPIErrorMapping(ref List<List<string>> errorMappng, ref Word.Document doc,string serviceFlow)
        {
            int tableNumber = 7;
            Word.Table ErrorTable = doc.Tables[tableNumber];
            int index = 2;
            foreach (List<string> errorRow in errorMappng)
            {
                if (serviceFlow == "Inbound")
                {
                    ErrorTable.Rows.Add(ErrorTable.Rows[1]);
                    //Return Code
                    Word.Cell cell = ErrorTable.Cell(index, 1);
                    cell.Range.Text = errorRow[1];
                    //Reason Code
                    cell = ErrorTable.Cell(index, 2);
                    cell.Range.Text = errorRow[2];
                    //Description
                    cell = ErrorTable.Cell(index, 3);
                    cell.Range.Text = errorRow[3];
                }
                else if (serviceFlow == "Outbound")
                {
                    ErrorTable.Rows.Add(ErrorTable.Rows[index]);
                    //Status Code
                    Word.Cell cell = ErrorTable.Cell(index, 1);
                    cell.Range.Text = errorRow[1];
                    //Description
                    cell = ErrorTable.Cell(index, 2);
                    cell.Range.Text = errorRow[3];
                }
                index++;
            }
            ErrorTable.Rows[index].Delete();
            Console.WriteLine("SPI Error Table Added");
        }
        public void addDTDErrorMapping(ref List<List<string>> errorMappng, ref Word.Document doc)
        {
            int tableNumber = 16;
            Word.Table ErrorTable = doc.Tables[tableNumber];
            int index = 2;
            foreach (List<string> errorRow in errorMappng)
            {
                ErrorTable.Rows.Add(ErrorTable.Rows[index]);
                
                Word.Cell cell;
                cell = ErrorTable.Cell(index, 1);
                cell.Range.Text = errorRow[0];
                cell = ErrorTable.Cell(index, 2);
                cell.Range.Text = errorRow[1];
                cell = ErrorTable.Cell(index, 3);
                cell.Range.Text = errorRow[2];
                cell = ErrorTable.Cell(index, 4);
                cell.Range.Text = errorRow[3];
                index++;
            }
            ErrorTable.Rows[index].Delete();
            Console.WriteLine("DTD Error Table Added");
        }
        private void updateDocument(ref Word.Document doc)
        {
            try
            {
                doc.Fields.Update();

                foreach (Word.Section section in doc.Sections)
                {
                    doc.Fields.Update();  // update each section

                    Word.HeadersFooters headers = section.Headers;  //Get all headers
                    foreach (Word.HeaderFooter header in headers)
                    {
                        Word.Fields fields = header.Range.Fields;
                        foreach (Word.Field field in fields)
                        {
                            field.Update();  // update all fields in headers
                        }
                    }

                    Word.HeadersFooters footers = section.Footers;  //Get all footers
                    foreach (Word.HeaderFooter footer in footers)
                    {
                        Word.Fields fields = footer.Range.Fields;
                        foreach (Word.Field field in fields)
                        {
                            field.Update();  //update all fields in footers
                        }
                    }
                }

                foreach (Word.TableOfContents tableOfContents in doc.TablesOfContents)
                {
                    tableOfContents.Update();
                }
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }
        public void copyTableData(List<int> oldTablesNumbers, int newTableNumber,ref Dictionary<string,string> dict, ref Dictionary<string, string> constantsDict, ref Word.Document newDoc)
        {
            string oldFilePath =constantsDict["Reading Old Documents Path"] + "\\INT-002-002-" + dict["OldServiceID"] + "-DTD " + dict["OldServiceSubject"] + ".docx";
            Console.WriteLine(oldFilePath);
            Word.Document oldDoc = wordApp.Documents.Open(oldFilePath);
            Word.Table newRequestTable = newDoc.Tables[newTableNumber];
            int index = 2;
            foreach (int k in oldTablesNumbers)
            {
                Console.WriteLine("Reading old Table Number "+k);
                Word.Table oldRequestTable = oldDoc.Tables[k];
                Word.Rows oldRequestRows = oldRequestTable.Rows;
                for (int i = 3; i <= oldRequestRows.Count - 1; i++)
                {
                    newRequestTable.Rows.Add(newRequestTable.Rows[index]);
                    Word.Row oldRow = oldRequestRows[i];
                    for (int j = 1; j <= oldRow.Cells.Count; j++)
                    {
                        Word.Cell cell1 = newRequestTable.Cell(index, j);
                        //Console.WriteLine(i + " " + j + " " + oldRequestTable.Cell(i, j).Range.Text);
                        cell1.Range.Text = oldRequestTable.Cell(i, j).Range.Text;
                    }
                    index++;
                }
            }
            newRequestTable.Rows[index].Delete();
            oldDoc.Close();
        }
        ~WordHandler()
        {
            wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            wordApp.Quit();
        }
    }
}
