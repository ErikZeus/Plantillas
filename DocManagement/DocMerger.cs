using System;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.Globalization;

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
        content = AccesoDatos.RegresaTablaSql("select    nombre + ' ' + segundo_nombre + ' ' + apellido + ' ' + segundo_apellido as NombreCompleto,direccion, apartado, case when isnull(s.cat_descr_catalogo,'') = '' then 'Sr./Sra.' else s.cat_descr_catalogo end as Titulo from clientes c left outer join seg_catalogo s on c.clas = s.cat_cod_catalogo and tab_cod_tabla = 'seg_clasificacion'  where cliente = " + _id);
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
        _Application wordApplication = new Application();

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
        DocMerger.Merge(arrFiles, strOutDoc, true, strOrgDoc);
    }

    public void CorrePlantilla1(string IdCliente, string archivo)
    {
        try
        {
            Plantillas Marco = new Plantillas();
            string CodigoCliente = AccesoDatos.RegresaCadena_1_ResultadoSql("Select cliente from CartasClientes where indice = " + IdCliente);
            Object oMissing = System.Reflection.Missing.Value;
            Object oTemplatePath = archivo + Marco.listado[0].UbicacionPlantilla;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc = new Microsoft.Office.Interop.Word.Document();
            wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);
            wordDoc.Activate();
            System.Data.DataTable content = InfoCliente(CodigoCliente);
            object cuerpo = AccesoDatos.RegresaCadena_1_ResultadoSql("Select cuerpo from CartasClientes where indice = " + IdCliente);
            string _poliza = AccesoDatos.RegresaCadena_1_ResultadoSql("Select poliza from CartasClientes where indice = " + IdCliente);
            string _asistente = AccesoDatos.RegresaCadena_1_ResultadoSql("Select asistente from CartasClientes where indice = " + IdCliente);
            string _fecha = AccesoDatos.RegresaCadena_1_ResultadoSql("Select Fecha from CartasClientes where indice = " + IdCliente);

            foreach (System.Data.DataRow rw in content.Rows)
            {

                DateTime FechaT = DateTime.Parse(_fecha.Substring(0,10));
                string _Fecha = "Guatemala " +  FechaT.Day.ToString() + " de " + CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(FechaT.Month) + " de " + FechaT.Year.ToString();
                FindAndReplace(wordApp, "{Fecha}", _Fecha);
                FindAndReplace(wordApp, "{Titulo}", rw["Titulo"]);
                FindAndReplace(wordApp, "{NombreCompleto}", rw["NombreCompleto"]);
                FindAndReplace(wordApp, "{Direccion}", rw["direccion"]);
                FindAndReplace(wordApp, "{Apartado}", rw["Apartado"]);
                FindAndReplace(wordApp, "{PolizasEjecutivo}", _poliza);
                FindAndReplace(wordApp, "{Asistente}", _asistente);

                break;
            }

            string[] arc = cuerpo.ToString().Split('.');
            int count = 1;
            int filas = arc.Length;

            string replace = "";

            foreach (string fl in arc)
            {

                if (count > 5)
                {
                    replace += fl;
                }
                else
                {
                    FindAndReplace(wordApp, "{Parrafo" + count.ToString() + "}", fl.Trim() + ".");
                }

                count += 1;
            }

            FindAndReplace(wordApp, "{Parrafo6}", replace);
            FindAndReplace(wordApp, "{Parrafo1}", "");
            FindAndReplace(wordApp, "{Parrafo2}", "");
            FindAndReplace(wordApp, "{Parrafo3}", "");
            FindAndReplace(wordApp, "{Parrafo4}", "");
            FindAndReplace(wordApp, "{Parrafo5}", "");
            FindAndReplace(wordApp, "{Parrafo6}", "");
            FindAndReplace(wordApp, "{Doc}","Doc corr." + IdCliente);

            wordDoc.SaveAs(archivo + Marco.listado[0].UbicacionMerge + CodigoCliente + ".docx");

            wordDoc.Close(true); // Close the Word Document.
            wordApp.Quit(true); // Close Word Application.


        }
        catch (Exception ex)
        {

            throw;

        }
    }


}

public class Plantilla
{
    public int id = 0;
    public string NombrePlantilla = "";
    public string UbicacionPlantilla = "";
    public string UbicacionMerge = "";
    public string Direccion = "";
    public Plantilla(int _id, string _NombrePlantilla, string _UbicacionPlantilla, string _UbicacionMerge, string _Direccion)
    {
        id = _id;
        NombrePlantilla = _NombrePlantilla;
        UbicacionPlantilla = _UbicacionPlantilla;
        UbicacionMerge = _UbicacionMerge;
        Direccion = _Direccion;
    }

}

public class Plantillas
{
    public int NumeroPlantillas = 6 ;

    public List<Plantilla> listado = new List<Plantilla>(new Plantilla[] {
        new Plantilla(1,"Carta envio ramo 9", @"\Files\OriginalDoc\Carta envio ramo 9.docx",@"\Files\Copies\","../Plantilla1.aspx"),
        new Plantilla(2,"Carta envio ramo 123", @"\Files\OriginalDoc\Carta envio ramo 123.docx",@"\Files\Copies\","../Plantilla2.aspx"),
        new Plantilla(3,"Carta envio todos excepto 9 y 123", @"\Files\OriginalDoc\Carta envio todos excepto 9 y 123.docx",@"\Files\Copies\","../Plantilla3.aspx"),
        new Plantilla(4,"Envio endosos ramo 9", @"\Files\OriginalDoc\Envio endosos ramo 9.docx",@"\Files\Copies\","../Plantilla4.aspx"),
         new Plantilla(5,"Envio endosos ramo 123", @"\Files\OriginalDoc\Envio endosos ramo 123.docx",@"\Files\Copies\","../Plantilla5.aspx"),
          new Plantilla(6,"Envio Endosos todos excepto 9 y 123", @"\Files\OriginalDoc\Envio Endosos todos excepto 9 y 123.docx",@"\Files\Copies\","../Plantilla6.aspx")
    });

}














