﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using Microsoft.Office.Interop.Word;



namespace DocManagement
{
    public partial class Plantilla6 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                string _id = Request.QueryString["Id"].ToString();
                string cliente = AccesoDatos.RegresaCadena_1_ResultadoSql("Select cliente from CartasClientes where indice = " + _id);

                DocMerger Doc = new DocMerger();
                string archivo = Server.MapPath(".");
                Doc.CorrePlantilla1(_id, archivo);
                Plantillas Marco = new Plantillas();
                
                Response.Redirect(archivo + Marco.listado[0].UbicacionMerge + cliente + ".docx");

            }
            catch (Exception es)
            {

                Helper.RegistrarEvento("Plantilla 1 " + es.Message);
            }
  


        }

 

    }
}