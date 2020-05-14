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
        
        public Word.Document generateSCI(Word.Application wordApp)
        {
            try
            {
                // Copy the template
                string DocFileName = Constants.TFSFolderPath + "\\SCI\\" + "INT-002-004-" + Constants.ServiceID + "-SCI " + Constants.Subject + ".docx";
                File.Copy(Constants.TemplatesFolderPath + "\\Templates\\" + Constants.ServiceFlow + "\\Template_SCI.docx", DocFileName, true);

                // wordApp.Visible = true;
                Word.Document doc = wordApp.Documents.Open(DocFileName);
                doc = wordApp.ActiveDocument;

                //setting the document properties
                doc.BuiltInDocumentProperties[Word.WdBuiltInProperty.wdPropertySubject].Value = Constants.Subject;

                doc.CustomDocumentProperties["Publish Date"].value = Constants.PublishDate;
                doc.CustomDocumentProperties["Review Date"].value = Constants.ReviewDate;
                doc.CustomDocumentProperties["Service Canonical Name"].value = Constants.ServiceCanonicalName;
                doc.CustomDocumentProperties["Service SubCategory"].value = Constants.ServiceSubCategory;
                doc.CustomDocumentProperties["Service ID"].value = Constants.ServiceID;

                //updating the document

                // Update properties
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
                return doc;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public Word.Document generateDTD(Word.Application wordApp)
        {
            try
            {
                // Copy the template
                string DocFileName = Constants.TFSFolderPath + "\\DTD\\" + "INT-002-004-" + Constants.ServiceID + "-DTD " + Constants.Subject + ".docx";
                File.Copy(Constants.TemplatesFolderPath + "\\Templates\\" + Constants.ServiceFlow + "\\Template_DTD.docx", DocFileName, true);

                // wordApp.Visible = true;
                Word.Document doc = wordApp.Documents.Open(DocFileName);
                doc = wordApp.ActiveDocument;

                //setting the document properties
                doc.BuiltInDocumentProperties[Word.WdBuiltInProperty.wdPropertySubject].Value = Constants.Subject;

                doc.CustomDocumentProperties["Publish Date"].value = Constants.PublishDate;
                doc.CustomDocumentProperties["Review Date"].value = Constants.ReviewDate;
                doc.CustomDocumentProperties["SSL Client Crypto Profile"].value = Constants.SSLClientCryptoProfile;
                doc.CustomDocumentProperties["XML Manager"].value = Constants.XMLManager;
                doc.CustomDocumentProperties["Service Canonical Name"].value = Constants.ServiceCanonicalName;
                doc.CustomDocumentProperties["Service SubCategory"].value = Constants.ServiceSubCategory;
                doc.CustomDocumentProperties["Service ID"].value = Constants.ServiceID;

                //updating the document

                // Update properties
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
                return doc;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public Word.Document generateSPI(Word.Application wordApp)
        {
            try
            {
                // Copy the template
                string DocFileName = Constants.TFSFolderPath + "\\SPI\\" + "INT-002-004-" + Constants.ServiceID + "-SPI " + Constants.Subject + ".docx";
                File.Copy(Constants.TemplatesFolderPath + "\\Templates\\" + Constants.ServiceFlow + "\\Template_SPI.docx", DocFileName, true);

                // wordApp.Visible = true;
                Word.Document doc = wordApp.Documents.Open(DocFileName);
                doc = wordApp.ActiveDocument;

                //setting the document properties
                doc.BuiltInDocumentProperties[Word.WdBuiltInProperty.wdPropertySubject].Value = Constants.Subject;

                doc.CustomDocumentProperties["Publish Date"].value = Constants.PublishDate;
                doc.CustomDocumentProperties["Review Date"].value = Constants.ReviewDate;
                doc.CustomDocumentProperties["Service Canonical Name"].value = Constants.ServiceCanonicalName;
                doc.CustomDocumentProperties["Service SubCategory"].value = Constants.ServiceSubCategory;
                doc.CustomDocumentProperties["Service ID"].value = Constants.ServiceID;

                //updating the document

                // Update properties
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
                return doc;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
