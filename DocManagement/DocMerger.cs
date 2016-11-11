using System;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using System.Drawing;
using System.Web;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.Office.Utils;

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

    public System.Data.DataTable InfoPoliza(string __poliza)
    {
        //select nombre + ' ' + segundo_nombre + ' ' + apellido + ' ' + segundo_apellido as NombreCompleto,nombre + ' ' + apellido as NombreApellido,nombre + ' ' + segundo_nombre as Nombres,apellido + ' ' + segundo_apellido as Apellidos,fechanac , cedula, ruc, direccion from clientes where cliente = 1
        System.Data.DataTable content = new System.Data.DataTable();
        content = AccesoDatos.RegresaTablaSql("select r.descr, p.poliza, ci.nombre, p.vigf as fecha_vencimiento_contrato, p.endoso, g.gst_nombre, g.gst_correo, cl.* from poliza p " +
  "    left outer join " +
  "   ramos r on p.ramo = r.ramo " +
  "    inner join " +
  "    ciaseg ci on p.cia = ci.cia " +
  "    inner join " +
  "    gestores g on p.gestor = g.gst_codigo_gestor " +
  "     left join " +
  "    (select  cliente, nombre + ' ' + segundo_nombre + ' ' + apellido + ' ' + segundo_apellido as NombreCompleto, " +
  "    direccion, isnull(apartado,'') as apartado, case when isnull(s.cat_descr_catalogo, '') = '' then 'Sr./Sra.' else s.cat_descr_catalogo end " +
  "    as Titulo from clientes c " +
  "    left outer join " +
  "    seg_catalogo s on c.clas = s.cat_cod_catalogo " +
  "    and tab_cod_tabla = 'seg_clasificacion') cl " +
  "    on p.cliente = cl.cliente " +
  "    where p.poliza = '" + __poliza + "' and p.tipo = 'poliza' and p.status != 'Cancelada' and secren = (select MAX(secren) from poliza where poliza ='" + __poliza + "' and tipo = 'poliza' and status != 'Cancelada' )");
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

            string CodigoCliente = AccesoDatos.RegresaCadena_1_ResultadoSql("Select cliente from CartasClientes where indice = " + IdCliente);
            Object oMissing = System.Reflection.Missing.Value;

            string carta = AccesoDatos.RegresaCadena_1_ResultadoSql("Select tipo_carta from CartasClientes where indice = " + IdCliente);
            try
            {

                Object oTemplatePath = Direccion(archivo, carta);
                RichEditDocumentServer wordDoc = new RichEditDocumentServer();
                wordDoc.LoadDocument(oTemplatePath.ToString());
                DocumentPosition pos = wordDoc.Document.CreatePosition(60);

                string CodigoImagen = AccesoDatos.RegresaCadena_1_ResultadoSql("Select codigo from CartasClientes where indice = " + IdCliente);

                Image DataImage = DevuelveImagen(CodigoImagen);
                string path = HttpContext.Current.Server.MapPath("~/Files/Copies/" + CodigoImagen + ".bmp");
                DataImage.Save(path);


                wordDoc.Document.InsertImage(pos, DataImage);


                string _poliza = "";
                string parrafo1, parrafo2, parrafo3 = "";
                string _tituloCarta = "";
                System.Data.DataTable poliza = AccesoDatos.RegresaTablaSql("Select * from CartasClientes where indice = " + IdCliente);
                foreach (System.Data.DataRow rw in poliza.Rows)
                {
                    DateTime FechaT = DateTime.Parse(DateTime.Now.ToShortDateString());
                    string _Fecha = "Guatemala, " + FechaT.Day.ToString() + " de " + CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(FechaT.Month) + " de " + FechaT.Year.ToString();
                    wordDoc.Document.ReplaceAll("{Fecha}", rw["Fecha"].ToString(), SearchOptions.WholeWord);
                    wordDoc.Document.ReplaceAll("{Titulo}", rw["Titulo"].ToString(), SearchOptions.WholeWord);
                    wordDoc.Document.ReplaceAll("{NombreCompleto}", rw["NombreCompleto"].ToString(), SearchOptions.WholeWord);
                    wordDoc.Document.ReplaceAll("{Nombre}", rw["NombreCompleto"].ToString(), SearchOptions.WholeWord);
                    wordDoc.Document.ReplaceAll("{Direccion}", rw["direccion"].ToString(), SearchOptions.WholeWord);
                    wordDoc.Document.ReplaceAll("{Apartado}", rw["Apartado"].ToString(), SearchOptions.WholeWord);
                    wordDoc.Document.ReplaceAll("{PolizasEjecutivo}", rw["poliza"].ToString(), SearchOptions.WholeWord);
                    wordDoc.Document.ReplaceAll("{Asistente}", rw["asistente"].ToString(), SearchOptions.WholeWord);


                    parrafo1 = rw["parrafo1"].ToString().Trim();
                    parrafo2 = rw["parrafo2"].ToString().Trim();
                    parrafo3 = rw["parrafo3"].ToString().Trim();
                    _tituloCarta = rw["TituloPagina"].ToString().Trim();
                    if (parrafo1 == "")
                    {
                        parrafo1 = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta = '" + rw["tipo_carta"].ToString() + "'");
                    }
                    if (parrafo2 == "")
                    {
                        parrafo2 = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta = '" + rw["tipo_carta"].ToString() + "'");
                    }
                    if (parrafo3 == "")
                    {
                        parrafo3 = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta = '" + rw["tipo_carta"].ToString() + "'");
                    }
                    if (_tituloCarta == "")
                    {
                        _tituloCarta = AccesoDatos.RegresaCadena_1_ResultadoSql("Select titulo from CartasClientes_Default where NombreCarta = '" + rw["tipo_carta"].ToString() + "'");

                    }

                    wordDoc.Document.ReplaceAll("{TITULO CARTA}", _tituloCarta, SearchOptions.WholeWord);
                    wordDoc.Document.ReplaceAll("{Parrafo1}", parrafo1, SearchOptions.WholeWord);
                    wordDoc.Document.ReplaceAll("{Parrafo2}", parrafo2, SearchOptions.WholeWord);
                    wordDoc.Document.ReplaceAll("{Parrafo3}", parrafo3, SearchOptions.WholeWord);
                    wordDoc.Document.ReplaceAll("{bien}", rw["bien"].ToString(), SearchOptions.WholeWord);
                    wordDoc.Document.ReplaceAll("{Poliza}", rw["poliza"].ToString(), SearchOptions.WholeWord);
                    _poliza = rw["poliza"].ToString();
                    wordDoc.Document.ReplaceAll("{aseguradora}", rw["aseguradora"].ToString(), SearchOptions.WholeWord);
                    string fechavence = rw["vence"].ToString();
                    if (fechavence.Length > 10)
                    {
                        wordDoc.Document.ReplaceAll("{vence}", rw["vence"].ToString().Substring(0, 10), SearchOptions.WholeWord);
                    }
                    else { wordDoc.Document.ReplaceAll("{vence}", fechavence, SearchOptions.WholeWord); }
                    wordDoc.Document.ReplaceAll("{endoso}", rw["endoso"].ToString(), SearchOptions.WholeWord);
                    wordDoc.Document.ReplaceAll("{Requerimiento}", rw["requerimiento"].ToString(), SearchOptions.WholeWord);
                    wordDoc.Document.ReplaceAll("{firma}", rw["firma"].ToString(), SearchOptions.WholeWord);
                    wordDoc.Document.ReplaceAll("{correo}", rw["correo"].ToString(), SearchOptions.WholeWord);
                    wordDoc.Document.ReplaceAll("{firma_administrativa}", rw["firma_administrativa"].ToString(), SearchOptions.WholeWord);
                    wordDoc.Document.ReplaceAll("{correo_administrativa}", rw["correo_administrativa"].ToString(), SearchOptions.WholeWord);


                    break;
                }
                wordDoc.Document.ReplaceAll("{Fecha}", "", SearchOptions.WholeWord);
                wordDoc.Document.ReplaceAll("{Titulo}", "", SearchOptions.WholeWord);
                wordDoc.Document.ReplaceAll("{NombreCompleto}", "", SearchOptions.WholeWord);
                wordDoc.Document.ReplaceAll("{Direccion}", "", SearchOptions.WholeWord);
                wordDoc.Document.ReplaceAll("{Apartado}", "", SearchOptions.WholeWord);
                wordDoc.Document.ReplaceAll("{PolizasEjecutivo}", "", SearchOptions.WholeWord);
                wordDoc.Document.ReplaceAll("{Asistente}", "", SearchOptions.WholeWord);
                wordDoc.Document.ReplaceAll("{Parrafo1}", "", SearchOptions.WholeWord);
                wordDoc.Document.ReplaceAll("{Parrafo2}", "", SearchOptions.WholeWord);
                wordDoc.Document.ReplaceAll("{Parrafo3}", "", SearchOptions.WholeWord);
                wordDoc.Document.ReplaceAll("{bien}", "", SearchOptions.WholeWord);
                wordDoc.Document.ReplaceAll("{PolizasEjecutivo}", "", SearchOptions.WholeWord);
                wordDoc.Document.ReplaceAll("{aseguradora}", "", SearchOptions.WholeWord);
                wordDoc.Document.ReplaceAll("{vence}", "", SearchOptions.WholeWord);
                wordDoc.Document.ReplaceAll("{endoso}", "", SearchOptions.WholeWord);
                wordDoc.Document.ReplaceAll("{Requerimiento}", "", SearchOptions.WholeWord);
                wordDoc.Document.ReplaceAll("{firma}", "", SearchOptions.WholeWord);
                wordDoc.Document.ReplaceAll("{correo}", "", SearchOptions.WholeWord);
                wordDoc.Document.ReplaceAll("{firma_administrativa}", "", SearchOptions.WholeWord);
                wordDoc.Document.ReplaceAll("{correo_administrativa}", "", SearchOptions.WholeWord);
                wordDoc.Document.ReplaceAll("{Nombre}", "", SearchOptions.WholeWord);
                wordDoc.Document.ReplaceAll("{TITULO CARTA}", "", SearchOptions.WholeWord);
                wordDoc.Document.ReplaceAll("{Doc}", "Doc corr." + IdCliente, SearchOptions.WholeWord);

                if (carta == "Envío de Documentos")
                {
                    int filas = Int32.Parse(AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from asegurado where poliza = '" + _poliza + "'"));
                    if (filas > 0)
                    {
                        TablaDoc(ref wordDoc, filas, _poliza);
                        wordDoc.Document.ReplaceAll("{grilla}", "", SearchOptions.WholeWord);
                    }
                    else
                    {
                        wordDoc.Document.ReplaceAll("{grilla}", "No tiene asegurados.", SearchOptions.WholeWord);
                    }
                }
                wordDoc.Document.Protect("UnitySecure");
                wordDoc.SaveDocument(path.Replace(".bmp", ".docx"), DocumentFormat.OpenXml);

                wordDoc.Dispose();
                DataImage.Dispose();
                System.GC.Collect();
            }
            catch (Exception es)
            {

                Helper.RegistrarEvento("Cargando el documento : " + es.Message);
            }


        }
        catch (Exception ex)
        {

            Helper.RegistrarEvento(ex.Message);

        }
    }

    public void TablaDoc(ref RichEditDocumentServer wordDoc, int filas, string _poliza)
    {
        DocumentPosition pos = wordDoc.Document.CreatePosition(1175);
        System.Data.DataTable content = AccesoDatos.RegresaTablaSql("select asegurado as nombre, certificado, clase as descripcion from asegurado where poliza = '" + _poliza + "' order by asegurado, certificado");
        DevExpress.XtraRichEdit.API.Native.Table table = wordDoc.Document.Tables.Add(pos, filas, 5);

        // Major adjustments
        table.TableLayout = TableLayoutType.Fixed;

        table.PreferredWidthType = WidthType.Fixed;
        table.PreferredWidth = Units.InchesToDocumentsF(11f);

        table.Rows[1].HeightType = HeightType.Exact;
        table.Rows[1].Height = Units.InchesToDocumentsF(0.25f);
        wordDoc.Document.InsertText(table[0, 0].Range.Start, "Nombre");
        wordDoc.Document.InsertText(table[0, 2].Range.Start, "Cert.");
        wordDoc.Document.InsertText(table[0, 4].Range.Start, "Descripción");
        table[0, 0].Style.Bold = true;
        table[0, 2].Style.Bold = true;
        table[0, 4].Style.Bold = true;
        table[0, 0].Style.FontSize = 12;
        table[0, 2].Style.FontSize = 12;
        table[0, 4].Style.FontSize = 12;

        table[0, 0].Borders.Bottom.LineColor = Color.White;
        table[0, 0].Borders.Top.LineColor = Color.White;
        table[0, 0].Borders.Left.LineColor = Color.White;
        table[0, 0].Borders.Right.LineColor = Color.White;
        table[0, 1].Borders.Bottom.LineColor = Color.White;
        table[0, 1].Borders.Top.LineColor = Color.White;
        table[0, 1].Borders.Left.LineColor = Color.White;
        table[0, 1].Borders.Right.LineColor = Color.White;
        table[0, 2].Borders.Bottom.LineColor = Color.White;
        table[0, 2].Borders.Top.LineColor = Color.White;
        table[0, 2].Borders.Left.LineColor = Color.White;
        table[0, 2].Borders.Right.LineColor = Color.White;
        table[0, 3].Borders.Bottom.LineColor = Color.White;
        table[0, 3].Borders.Top.LineColor = Color.White;
        table[0, 3].Borders.Left.LineColor = Color.White;
        table[0, 3].Borders.Right.LineColor = Color.White;
        table[0, 4].Borders.Bottom.LineColor = Color.White;
        table[0, 4].Borders.Top.LineColor = Color.White;
        table[0, 4].Borders.Left.LineColor = Color.White;
        table[0, 4].Borders.Right.LineColor = Color.White;
        // table[0, 0]..Right.LineColor = Color.White;
        int tupla = 1;
        // Additional adjustments
        //DocumentPosition pos = wordDoc.Document.CreatePosition(60);
        foreach (System.Data.DataRow rw in content.Rows)
        {
            try
            {
                table.Rows[tupla].Cells[0].BackgroundColor = Color.Yellow;
                table.Rows[tupla].Cells[2].BackgroundColor = Color.Yellow;
                table.Rows[tupla].Cells[4].BackgroundColor = Color.Yellow;
                table[tupla, 0].PreferredWidthType = WidthType.Fixed;
                table[tupla, 0].PreferredWidth = Units.InchesToDocumentsF(11f);
                table[tupla, 1].PreferredWidthType = WidthType.Fixed;
                table[tupla, 1].PreferredWidth = Units.InchesToDocumentsF(1f);
                table[tupla, 1].BackgroundColor = Color.White;
                table[tupla, 2].PreferredWidthType = WidthType.Fixed;
                table[tupla, 2].PreferredWidth = Units.InchesToDocumentsF(1f);
                table[tupla, 3].PreferredWidthType = WidthType.Fixed;
                table[tupla, 3].PreferredWidth = Units.InchesToDocumentsF(1f);
                table[tupla, 3].BackgroundColor = Color.White;
                table[tupla, 4].PreferredWidthType = WidthType.Fixed;
                table[tupla, 4].PreferredWidth = Units.InchesToDocumentsF(2f);


                wordDoc.Document.InsertText(table[tupla, 0].Range.Start, rw["nombre"].ToString());
                wordDoc.Document.InsertText(table[tupla, 1].Range.Start, "");
                wordDoc.Document.InsertText(table[tupla, 2].Range.Start, rw["certificado"].ToString());
                wordDoc.Document.InsertText(table[tupla, 3].Range.Start, "");
                wordDoc.Document.InsertText(table[tupla, 4].Range.Start, rw["descripcion"].ToString());

                tupla += 1;
            }
            catch (Exception we)
            {

                Helper.RegistrarEvento("Cargando wordDoc : " + tupla.ToString() + " - " + we.Message);
            }

        }

        table[1, 1].LeftPadding = 0;


    }
    private string Direccion(string archivo, string carta)
    {

        Plantillas Marco = new Plantillas();

        if (carta == "Carta Envío ramo 9")
        {
            return archivo + Marco.listado[0].UbicacionPlantilla;
        }
        if (carta == "Carta Envío ramo 123")
        {
            return archivo + Marco.listado[1].UbicacionPlantilla;
        }
        if (carta == "Carta Envío todos excepto 9 y 123")
        {
            return archivo + Marco.listado[2].UbicacionPlantilla;
        }
        if (carta == "Envío endosos ramo 9")
        {
            return archivo + Marco.listado[3].UbicacionPlantilla;
        }
        if (carta == "Envío endosos ramo 123")
        {
            return archivo + Marco.listado[4].UbicacionPlantilla;
        }
        if (carta == "Envío Endosos todos excepto 9 y 123")
        {
            return archivo + Marco.listado[5].UbicacionPlantilla;
        }
        if (carta == "Envío de Documentos Incendio")
        {
            return archivo + Marco.listado[6].UbicacionPlantilla;
        }
        if (carta == "Envío de Documentos Equipo Electronico")
        {
            return archivo + Marco.listado[7].UbicacionPlantilla;
        }
        if (carta == "Envío de Documentos en blanco")
        {
            return archivo + Marco.listado[8].UbicacionPlantilla;
        }
        if (carta == "Envío de Documentos Autos")
        {
            return archivo + Marco.listado[9].UbicacionPlantilla;
        }
        if (carta == "Envío de Cartapacio")
        {
            return archivo + Marco.listado[10].UbicacionPlantilla;
        }
        if (carta == "Aviso de Renovación con Condiciones Autos")
        {
            return archivo + Marco.listado[11].UbicacionPlantilla;
        }
        if (carta == "Envíos Varios en Blanco")
        {
            return archivo + Marco.listado[12].UbicacionPlantilla;
        }
        if (carta == "Aviso de Renovación con Condiciones Incendio")
        {
            return archivo + Marco.listado[13].UbicacionPlantilla;
        }
        if (carta == "Envío de Renovación Autos y Daños")
        {
            return archivo + Marco.listado[14].UbicacionPlantilla;
        }
        if (carta == "Envío de Endoso Autos y Daños")
        {
            return archivo + Marco.listado[15].UbicacionPlantilla;
        }
        if (carta == "Envío de Renovación  Nueva Autos  Daños")
        {
            return archivo + Marco.listado[16].UbicacionPlantilla;
        }
        if (carta == "Envío de Documentos")
        {
            return archivo + Marco.listado[17].UbicacionPlantilla;
        }
        if (carta == "Envío de liquidación de reclamo")
        {
            return archivo + Marco.listado[18].UbicacionPlantilla;
        }
        if (carta == "Envío de Cheque y Liquidación de reclamo")
        {
            return archivo + Marco.listado[19].UbicacionPlantilla;
        }
        if (carta == "Envío de Facturación")
        {
            return archivo + Marco.listado[20].UbicacionPlantilla;
        }
        if (carta == "Envío de Endoso Vida y Gastos Médicos")
        {
            return archivo + Marco.listado[21].UbicacionPlantilla;
        }
        if (carta == "Envío de Renovación Vida y Gastos Médicos")
        {
            return archivo + Marco.listado[22].UbicacionPlantilla;
        }

        return "";
    }

    public System.Drawing.Image DevuelveImagen(string codigoimagen)
    {
        //Read in the parameters
        string strData = codigoimagen;
        int imageHeight = Convert.ToInt32("15");
        int imageWidth = Convert.ToInt32("200");
        string Forecolor = "";
        string Backcolor = "";
        bool bIncludeLabel = false;
        string strImageFormat = "BMP";
        string strAlignment = "c";

        BarcodeLib.TYPE type = BarcodeLib.TYPE.UNSPECIFIED;
        type = BarcodeLib.TYPE.CODE128;

        System.Drawing.Image barcodeImage = null;
        try
        {
            BarcodeLib.Barcode b = new BarcodeLib.Barcode();
            if (type != BarcodeLib.TYPE.UNSPECIFIED)
            {
                b.IncludeLabel = bIncludeLabel;

                //alignment
                switch (strAlignment.ToLower().Trim())
                {
                    case "c":
                        b.Alignment = BarcodeLib.AlignmentPositions.CENTER;
                        break;
                    case "r":
                        b.Alignment = BarcodeLib.AlignmentPositions.RIGHT;
                        break;
                    case "l":
                        b.Alignment = BarcodeLib.AlignmentPositions.LEFT;
                        break;
                    default:
                        b.Alignment = BarcodeLib.AlignmentPositions.CENTER;
                        break;
                }//switch

                if (Forecolor.Trim() == "" && Forecolor.Trim().Length != 6)
                    Forecolor = "000000";
                if (Backcolor.Trim() == "" && Backcolor.Trim().Length != 6)
                    Backcolor = "FFFFFF";

                //===== Encoding performed here =====
                barcodeImage = b.Encode(type, strData.Trim(), System.Drawing.ColorTranslator.FromHtml("#" + Forecolor), System.Drawing.ColorTranslator.FromHtml("#" + Backcolor), imageWidth, imageHeight);
                //===================================


            }//if
        }//try
        catch (Exception ex)
        {
            Helper.RegistrarEvento("Haciendo el codigo b image : " + ex.Message);
            //TODO: find a way to return this to display the encoding error message
        }//catch
        finally
        {

        }//finally
        return barcodeImage;
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
    public int NumeroPlantillas = 23;

    public List<Plantilla> listado = new List<Plantilla>(new Plantilla[] {
       new Plantilla(1,"Carta Envío ramo 9", @"\Files\OriginalDoc\Carta Envío ramo 9.docx",@"\Files\Copies\","../Plantilla1.aspx"),
       new Plantilla(2,"Carta Envío ramo 123", @"\Files\OriginalDoc\Carta Envío ramo 123.docx",@"\Files\Copies\","../Plantilla2.aspx"),
       new Plantilla(3,"Carta Envío todos excepto 9 y 123", @"\Files\OriginalDoc\Carta Envío todos excepto 9 y 123.docx",@"\Files\Copies\","../Plantilla3.aspx"),
       new Plantilla(4,"Envío endosos ramo 9", @"\Files\OriginalDoc\Envío endosos ramo 9.docx",@"\Files\Copies\","../Plantilla4.aspx"),
       new Plantilla(5,"Envío endosos ramo 123", @"\Files\OriginalDoc\Envío endosos ramo 123.docx",@"\Files\Copies\","../Plantilla5.aspx"),
       new Plantilla(6,"Envío Endosos todos excepto 9 y 123", @"\Files\OriginalDoc\Envío Endosos todos excepto 9 y 123.docx",@"\Files\Copies\","../Plantilla6.aspx"),
       new Plantilla(7,"Envío de Documentos Incendio", @"\Files\OriginalDoc\Envío de Documentos Incendio.docx",@"\Files\Copies\","../Plantilla7.aspx"),
       new Plantilla(8,"Envío de Documentos Equipo Electronico", @"\Files\OriginalDoc\Envío de Documentos Equipo Electronico.docx",@"\Files\Copies\","../Plantilla8.aspx"),
       new Plantilla(9,"Envío de Documentos en blanco", @"\Files\OriginalDoc\Envío de Documentos en blanco.docx",@"\Files\Copies\","../Plantilla9.aspx"),
       new Plantilla(10,"Envío de Documentos Autos", @"\Files\OriginalDoc\Envío de Documentos Autos.docx",@"\Files\Copies\","../Plantilla10.aspx"),
       new Plantilla(11,"Envío de Cartapacio", @"\Files\OriginalDoc\Envío de Cartapacio.docx",@"\Files\Copies\","../Plantilla11.aspx"),
       new Plantilla(12,"Aviso de Renovación con Condiciones Autos", @"\Files\OriginalDoc\Aviso de Renovación con Condiciones Autos.docx",@"\Files\Copies\","../Plantilla12.aspx"),
       new Plantilla(13,"Envíos Varios en Blanco", @"\Files\OriginalDoc\Envíos Varios en Blanco.docx",@"\Files\Copies\","../Plantilla13.aspx"),
       new Plantilla(14,"Aviso de Renovación con Condiciones Incendio", @"\Files\OriginalDoc\Aviso de Renovación con Condiciones Incendio.docx",@"\Files\Copies\","../Plantilla14.aspx"),
       new Plantilla(15,"Envío de Renovación Autos y Daños", @"\Files\OriginalDoc\Envío de Renovación Autos y Daños.docx",@"\Files\Copies\","../Plantilla15.aspx"),
       new Plantilla(16,"Envío de Endoso Autos y Daños", @"\Files\OriginalDoc\Envío de Endoso Autos y Daños.docx",@"\Files\Copies\","../Plantilla16.aspx"),
       new Plantilla(17,"Envío de Renovación  Nueva Autos  Daños", @"\Files\OriginalDoc\Envío de Renovación  Nueva Autos  Daños.docx",@"\Files\Copies\","../Plantilla17.aspx"),
       new Plantilla(18,"Envío de Documentos", @"\Files\OriginalDoc\Envío de Documentos.docx",@"\Files\Copies\","../Plantilla18.aspx"),
       new Plantilla(19,"Envío de liquidación de reclamo", @"\Files\OriginalDoc\Envío de liquidación de reclamo.docx",@"\Files\Copies\","../Plantilla19.aspx"),
       new Plantilla(20,"Envío de Cheque y Liquidación de reclamo", @"\Files\OriginalDoc\Envío de Cheque y Liquidación de reclamo.docx",@"\Files\Copies\","../Plantilla20.aspx"),
       new Plantilla(21,"Envío de Facturación", @"\Files\OriginalDoc\Envío de Facturación.docx",@"\Files\Copies\","../Plantilla21.aspx"),
       new Plantilla(22,"Envío de Endoso Vida y Gastos Médicos", @"\Files\OriginalDoc\Envío de Endoso Vida y Gastos Médicos.docx",@"\Files\Copies\","../Plantilla22.aspx"),
       new Plantilla(23,"Envío de Renovación Vida y Gastos Médicos", @"\Files\OriginalDoc\Envío de Renovación Vida y Gastos Médicos.docx",@"\Files\Copies\","../Plantilla23.aspx")
    });

}














