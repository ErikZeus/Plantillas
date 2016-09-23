using System;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;

/// <summary>
/// Simple class to get a merged document based on a baseline document
/// and a set of copies
/// </summary>
public class DocMerger
  {
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

    /// <summary>
    /// A function that merges Microsoft Word Documents that uses a template specified by the user
    /// </summary>
    /// <param name="filesToMerge">An array of files that we want to merge</param>
    /// <param name="outputFilename">The filename of the merged document</param>
    /// <param name="insertPageBreaks">Set to true if you want to have page breaks inserted after each document</param>
    /// <param name="documentTemplate">The word document you want to use to serve as the template</param>
    public static void Merge(string[] filesToMerge, string outputFilename, bool insertPageBreaks, string documentTemplate)
    {
        object defaultTemplate = documentTemplate;
        object missing = System.Type.Missing;
        object pageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
        object outputFile = outputFilename;

        // Create  a new Word application
         _Application wordApplication = new  Application();

        try
        {
            // Create a new file based on our template
            _Document wordDocument = wordApplication.Documents.Add(
                                          ref defaultTemplate
                                        , ref missing
                                        , ref missing
                                        , ref missing);

            // Make a Word selection object.
            Selection selection = wordApplication.Selection;

            // Loop thru each of the Word documents
            foreach (string file in filesToMerge)
            {
                // Insert the files to our template
                selection.InsertFile(
                                            file
                                        , ref missing
                                        , ref missing
                                        , ref missing
                                        , ref missing);

                //Do we want page breaks added after each documents?
                if (insertPageBreaks)
                {
                    selection.InsertBreak(ref pageBreak);
                }
            }

            // Save the document to it's output file.
            wordDocument.SaveAs(
                            ref outputFile
                        , ref missing
                        , ref missing
                        , ref missing
                        , ref missing
                        , ref missing
                        , ref missing
                        , ref missing
                        , ref missing
                        , ref missing
                        , ref missing
                        , ref missing
                        , ref missing
                        , ref missing
                        , ref missing
                        , ref missing);

            // Clean up!
            wordDocument = null;
        }
        catch (Exception ex)
        {
            //I didn't include a default error handler so i'm just throwing the error
            throw ex;
        }
        finally
        {
            // Finally, Close our Word application
            wordApplication.Quit(ref missing, ref missing, ref missing);
        }
    }
 


    /// <summary>
    /// Merge a document with a set of copies
    /// </summary>
    /// <param name="strOrgDoc">
    /// Original file name
    /// </param>
    /// <param name="strCopyFolder">
    /// Folder in which the copies are located
    /// </param>
    /// <param name="strOutDoc">
    /// The result filename
    /// </param>
    public void Merge(string strOrgDoc, string strCopyFolder, string strOutDoc)
    {
 
      string[] arrFiles = Directory.GetFiles(strCopyFolder);
       DocMerger.Merge(arrFiles, strOutDoc,true, strOrgDoc);
    }

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
        

        }
        catch (Exception ex)
        {
           
        }
    }


}


public class Plantilla
{
    public string  NombrePlantilla = "";
    public string UbicacionPlantilla = "";
    public int id = 0;
    public string UbicacionMerge = "";

    public Plantilla(string _NombrePlantilla, string _UbicacionPlantilla, int _id, string _UbicacionMerge) {
        NombrePlantilla = _NombrePlantilla;
        UbicacionPlantilla = _UbicacionPlantilla;
        id = _id;
        UbicacionMerge = _UbicacionMerge;
    }
 
}

public class Plantillas
{
    public int NumeroPlantillas = 3;

    public List<Plantilla> listado = new List<Plantilla>(new Plantilla[] {
        new Plantilla("Plantilla Carta1","files/OriginalDoc/CartaPremio1.docx",1,"files/Copies/CartaPremio1.docx"),
        new Plantilla("Plantilla Carta2","files/OriginalDoc/CartaPremio2.docx",2,"files/Copies/CartaPremio2.docx"),
        new Plantilla("Plantilla Carta3","files/OriginalDoc/CartaPremio3.docx",3,"files/Copies/CartaPremio2.docx") });

}

public class Helper
{

    public static void RegistrarEvento(string EventoInfo)
    {
        string YearFolder = System.DateTime.Now.Year.ToString();
        string LogFolderDir = Path.Combine(System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + "/Logs", YearFolder);
        string LogFileName = Path.Combine(LogFolderDir, System.DateTime.Now.ToString("yyyyMMdd") + ".txt");
        FileStream fs = default(FileStream);
        StreamWriter sw = default(StreamWriter);

        try
        {
            if (!Directory.Exists(LogFolderDir))
            {
                Directory.CreateDirectory(LogFolderDir);
            }

            if (!File.Exists(LogFileName))
            {
                fs = File.Create(LogFileName);
            }
            else
            {
                fs = new FileStream(LogFileName, FileMode.Append, FileAccess.Write);
            }

            sw = new StreamWriter(fs);

            sw.WriteLine("[{0}] - {1}", System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), EventoInfo);

            sw.Close();
            fs.Close();

        }
        catch (Exception ex)
        {
        }
    }

}







