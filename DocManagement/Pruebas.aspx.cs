using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;

using Microsoft.Office.Interop.Word;

namespace DocManagement
{
    public partial class Pruebas : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
             
        }
        //Create document method
        private void Plantilla1(string IdCliente)
        {
            try
            {
 
            Object oMissing = System.Reflection.Missing.Value;
            Object oTemplatePath = @"C:\Proyecto Administrador Doc Word\DocManagement\Files\OriginalDoc\CartaPremio1.docx";
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc = new Microsoft.Office.Interop.Word.Document();
            wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);
            wordDoc.Activate();
            FindAndReplace(wordApp, "{id}", IdCliente);
            wordDoc.SaveAs(@"C:\Proyecto Administrador Doc Word\DocManagement\Files\OriginalDoc\CartaPremioNuevo.docx");
            Label1.Text = "Document created successfully !";

            }
            catch (Exception ex)
            {
                Label1.Text = ex.Message;
            }
        }
        private void FindAndReplace(Microsoft.Office.Interop.Word.Application doc, object findText, object replaceWithText)
        {
            //options
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            doc.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }

    }
}