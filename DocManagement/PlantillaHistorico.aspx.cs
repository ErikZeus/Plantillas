using System;
using System.Web.UI.WebControls;
using System.Web.UI;


namespace DocManagement
{
    public partial class PlantillaHistorico : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                string id = Request.QueryString["id"];
                string cliente = AccesoDatos.RegresaCadena_1_ResultadoSql("Select cliente from CartasClientes where indice = " + id);
                this.lblCliente.Text = cliente;

                    Plantillas Marco = new Plantillas();
                    TableCell c = new TableCell();
                    TableRow r = new TableRow();
                    HyperLink hyp = new HyperLink();
                    hyp.Text = Marco.listado[3].NombrePlantilla;
                    hyp.NavigateUrl = Marco.listado[3].Direccion + "?id=" + id;
                    c.Controls.Add(hyp);
                    r.Cells.Add(c);

                    Table1.Rows.Add(r);
        
            }
            catch (Exception)
            {
 
            }



        }

 






    }
}