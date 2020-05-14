﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Visio = Microsoft.Office.Interop.Visio;

namespace OfficeHelper
{
    class VisioHandler
    {
        public void addSequanceDiagram(ref Word.Document doc)
        {
            try
            {
                Visio.Application visioApp = new Visio.Application();
                visioApp.Visible = false;
                Visio.Document sequenceVisio;
                string sequenceDiagramPath = Constants.TemplatesFolderPath + "\\Templates\\" + Constants.ServiceFlow + "\\ServiceSequanceDiagram.vsdx";
                sequenceVisio = visioApp.Documents.Open(sequenceDiagramPath);
                Visio.Page sequencePage = sequenceVisio.Pages[1];
                foreach (Visio.Shape shp in sequencePage.Shapes)
                {
                    if (shp.Name != "Watermark Title" && shp.Text != "")
                    {
                        shp.Text = shp.Text.Replace("ServiceCanonicalName",Constants.ServiceCanonicalName);
                        shp.Text = shp.Text.Replace("Subject", Constants.Subject);
                        shp.Text = shp.Text.Replace("Backend", Constants.BackendName);
                    }
                }

                Visio.Selection sequenceDiagram = sequencePage.CreateSelection(Visio.VisSelectionTypes.visSelTypeAll);

                sequenceDiagram.Copy();
                //rollback
                foreach (Visio.Shape shp in sequencePage.Shapes)
                {
                    if (shp.Name != "Watermark Title" && shp.Text != "")
                    {
                        shp.Text = shp.Text.Replace(Constants.ServiceCanonicalName, "ServiceCanonicalName");
                        shp.Text = shp.Text.Replace(Constants.Subject, "Subject");
                        shp.Text = shp.Text.Replace(Constants.BackendName, "Backend");
                    }
                }
                sequenceVisio.Saved = true;


                Word.Bookmark sequenceDiagBM = doc.Bookmarks["SequenceDiagram"];
                Word.Range sequenceDiagRng = sequenceDiagBM.Range;
                sequenceDiagRng.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting);

                

                visioApp.Quit();
            }
            catch(Exception ex)
            {
                throw ex;
            }


        }

        public void addRequestFlowDiagram(ref Word.Document doc)
        {
            try
            {
                Visio.Application visioApp = new Visio.Application();
                visioApp.Visible = false;
                Visio.Document requestVisio;
                string RequestDiagramPath = Constants.TemplatesFolderPath + "\\Templates\\" + Constants.ServiceFlow + "\\ServiceRequestFlow.vsdx";
                requestVisio = visioApp.Documents.Open(RequestDiagramPath);
                Visio.Page requestPage = requestVisio.Pages[1];
                foreach (Visio.Shape shp in requestPage.Shapes)
                {
                    if (shp.Name != "Watermark Title" && shp.Text != "")
                    {
                        shp.Text = shp.Text.Replace("ServiceCanonicalName", Constants.ServiceCanonicalName);
                    }
                }

                Visio.Selection requestDiagram = requestPage.CreateSelection(Visio.VisSelectionTypes.visSelTypeAll);

                requestDiagram.Copy();
                //rollback
                foreach (Visio.Shape shp in requestPage.Shapes)
                {
                    if (shp.Name != "Watermark Title" && shp.Text != "")
                    {
                        shp.Text = shp.Text.Replace(Constants.ServiceCanonicalName, "ServiceCanonicalName");
                    }
                }
                requestVisio.Saved = true;


                Word.Bookmark requestDiagBM = doc.Bookmarks["RequestFlow"];
                Word.Range requestDiagRng = requestDiagBM.Range;
                requestDiagRng.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting);

                visioApp.Quit();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void addResponseFlowDiagram(ref Word.Document doc)
        {
            try
            {
                Visio.Application visioApp = new Visio.Application();
                visioApp.Visible = false;
                Visio.Document ResponseVisio;
                string ResponseDiagramPath = Constants.TemplatesFolderPath + "\\Templates\\" + Constants.ServiceFlow + "\\ServiceResponseFlow.vsdx";
                ResponseVisio = visioApp.Documents.Open(ResponseDiagramPath);
                Visio.Page ResponsePage = ResponseVisio.Pages[1];
                foreach (Visio.Shape shp in ResponsePage.Shapes)
                {
                    if (shp.Name != "Watermark Title" && shp.Text != "")
                    {
                        shp.Text = shp.Text.Replace("ServiceCanonicalName", Constants.ServiceCanonicalName);
                    }
                }

                Visio.Selection requestDiagram = ResponsePage.CreateSelection(Visio.VisSelectionTypes.visSelTypeAll);

                requestDiagram.Copy();
                //rollback
                foreach (Visio.Shape shp in ResponsePage.Shapes)
                {
                    if (shp.Name != "Watermark Title" && shp.Text != "")
                    {
                        shp.Text = shp.Text.Replace(Constants.ServiceCanonicalName, "ServiceCanonicalName");
                    }
                }
                ResponseVisio.Saved = true;


                Word.Bookmark ResponseDiagBM = doc.Bookmarks["ResponseFlow"];
                Word.Range ResponseDiagRng = ResponseDiagBM.Range;
                ResponseDiagRng.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting);

                visioApp.Quit();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
