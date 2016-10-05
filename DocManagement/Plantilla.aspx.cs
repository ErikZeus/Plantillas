﻿using System;
using System.Web.UI.WebControls;
using System.Web.UI;


namespace DocManagement
{
    public partial class Plantilla : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                string id = Request.QueryString["id"];
                string cliente = AccesoDatos.RegresaCadena_1_ResultadoSql("Select cliente from CartasClientes where indice = " + id);
                this.lblCliente.Text = cliente;

                Plantillas Marco = new Plantillas();
                                

                for (int i = 0; i < Marco.NumeroPlantillas; i++)
                {
                    TableCell c = new TableCell();
                    TableRow r = new TableRow();
                    HyperLink hyp = new HyperLink();
                    hyp.Text = Marco.listado[i].NombrePlantilla;
                    hyp.NavigateUrl = Marco.listado[i].Direccion + "?id=" + id;
                    c.Controls.Add(hyp);
                    r.Cells.Add(c);

                    Table1.Rows.Add(r);
                }
            }
            catch (Exception)
            {
 
            }



        }

 






    }
}