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
        
        public Word.Document generateDocument(ref Dictionary<string, string> dict,string documentName)
        {
            try
            {
                // Copy the template
                string DocFileName = dict["TFSFolderPath"] + "\\"+documentName+"\\" + "INT-002-004-" + dict["ServiceID"] + "-"+documentName+" " + dict["Subject"] + ".docx";
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
        public void copyTableData(int oldTableNumber,int newTableNumber,ref Dictionary<string,string> dict, ref Word.Document newDoc)
        {
            string oldFilePath = Constants.OldDTDFilesPath + "INT-002-002-" + dict["OldServiceID"] + "-DTD " + dict["OldServiceSubject"] + ".docx";
            Console.WriteLine(oldFilePath);
            Word.Document oldDoc = wordApp.Documents.Open(oldFilePath);
            Word.Table oldRequestTable = oldDoc.Tables[oldTableNumber];
            Word.Rows oldRequestRows = oldRequestTable.Rows;

            Word.Table newRequestTable = newDoc.Tables[newTableNumber];
            Word.Rows newRequestRows = newRequestTable.Rows;

            int index = 2;
            for (int i = 3; i <= oldRequestRows.Count - 1; i++)
            {
                newRequestTable.Rows.Add(newRequestTable.Rows[index]);
                for (int j = 1; j <= 3; j++)
                {
                    Word.Cell cell1 = newRequestTable.Cell(index, j);
                    Console.WriteLine(oldRequestTable.Cell(i, j).Range.Text);
                    cell1.Range.Text = oldRequestTable.Cell(i, j).Range.Text;
                }
                index++;
            }
            newRequestTable.Rows[index].Delete();
        }
        ~WordHandler()
        {
            wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            wordApp.Quit();
        }
    }
}
