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

    public System.Data.DataTable InfoCliente(string _id)
    {
        //select nombre + ' ' + segundo_nombre + ' ' + apellido + ' ' + segundo_apellido as NombreCompleto,nombre + ' ' + apellido as NombreApellido,nombre + ' ' + segundo_nombre as Nombres,apellido + ' ' + segundo_apellido as Apellidos,fechanac , cedula, ruc, direccion from clientes where cliente = 1
       System.Data.DataTable content = new System.Data.DataTable();
       content = AccesoDatos.RegresaTablaSql("select nombre + ' ' + segundo_nombre + ' ' + apellido + ' ' + segundo_apellido as NombreCompleto,nombre + ' ' + apellido as NombreApellido,nombre + ' ' + segundo_nombre as Nombres,apellido + ' ' + segundo_apellido as Apellidos,fechanac , cedula, ruc, direccion from clientes where cliente = " + _id);
       return content;
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

    public  void CorrePlantilla1(string IdCliente)
    {
        try
        {
            Plantillas Marco = new Plantillas();

            Object oMissing = System.Reflection.Missing.Value;
            Object oTemplatePath = Marco.listado[0].UbicacionPlantilla;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc = new Microsoft.Office.Interop.Word.Document();
            wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);
            wordDoc.Activate();
            System.Data.DataTable content = InfoCliente(IdCliente);
            foreach (System.Data.DataRow rw in content.Rows)
            {
                FindAndReplace(wordApp, "{id}", rw["Nombres"]);

                break;
            }
           
            wordDoc.SaveAs(Marco.listado[0].UbicacionMerge + IdCliente + ".docx");
        
        }
        catch (Exception ex)
        {
           
        }
    }


}


public class Plantilla
{
    public int id = 0;
    public string  NombrePlantilla = "";
    public string UbicacionPlantilla = "";
    public string UbicacionMerge = "";
    public string Direccion = "";
    public Plantilla(int _id, string _NombrePlantilla, string _UbicacionPlantilla,  string _UbicacionMerge, string _Direccion) {
        id = _id;
        NombrePlantilla = _NombrePlantilla;
        UbicacionPlantilla = _UbicacionPlantilla;
        UbicacionMerge = _UbicacionMerge;
        Direccion = _Direccion;
    }
 
}

public class Plantillas
{
    public int NumeroPlantillas = 3;

    public List<Plantilla> listado = new List<Plantilla>(new Plantilla[] {
        new Plantilla(1,"Plantilla Carta1", @"C:\Proyecto Administrador Doc Word\DocManagement\Files\OriginalDoc\CartaPremio1.docx",@"C:\Proyecto Administrador Doc Word\DocManagement\Files\Copies\","../Plantilla1.aspx"),
        new Plantilla(2,"Plantilla Carta2", @"C:\Proyecto Administrador Doc Word\DocManagement\Files\OriginalDoc\CartaPremio2.docx",@"C:\Proyecto Administrador Doc Word\DocManagement\Files\Copies\","../Plantilla2.aspx"),
        new Plantilla(3,"Plantilla Carta3", @"C:\Proyecto Administrador Doc Word\DocManagement\Files\OriginalDoc\CartaPremio3.docx",@"C:\Proyecto Administrador Doc Word\DocManagement\Files\Copies\","../Plantilla3.aspx" ) });
}












