using System;
using System.IO;



namespace DocManagement
{
    public partial class Plantilla1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                string _id = Request.QueryString["id"].ToString();
  
                DocMerger Doc = new DocMerger();
                string archivo = Server.MapPath(".");
                Doc.CorrePlantilla1(_id, archivo);
                Plantillas Marco = new Plantillas();
                string address = archivo + Marco.listado[0].UbicacionMerge + _id + ".docx";


                MemoryStream memoryStream = new MemoryStream();
                using (Stream input = File.OpenRead(address))
                {
                    byte[] buffer = new byte[32 * 1024]; // 32K buffer for example
                    int bytesRead;
                    while ((bytesRead = input.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        memoryStream.Write(buffer, 0, bytesRead);
                    }
                }
                memoryStream.Position = 0;
                string filenamesend = "attachment; filename="+ _id.ToString() +".docx";

                Response.Clear();
                Response.ContentType = "Application/msword";
                Response.AddHeader("Content-Disposition", filenamesend);
                Response.BinaryWrite(memoryStream.ToArray());
                // myMemoryStream.WriteTo(Response.OutputStream); //works too
                Response.Flush();
                Response.Close();
                Response.End();


                

            }
            catch (Exception es)
            {

              //  Helper.RegistrarEvento("Plantilla 1 " + es.Message);
            }



        }


    }
}