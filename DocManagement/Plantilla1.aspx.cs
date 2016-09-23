using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;

using Microsoft.Office.Interop.Word;

namespace DocManagement
{
    public partial class Plantilla1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                string _id = Request.QueryString["Id"].ToString();
                DocMerger Doc = new DocMerger();
                Doc.CorrePlantilla1(_id);
                Plantillas Marco = new Plantillas();

                Response.Redirect(Marco.listado[0].UbicacionMerge + _id + ".docx");
            }
            catch (Exception es)
            {

                Helper.RegistrarEvento("Plantilla 1 " + es.Message);
            }
  


        }

 

    }
}