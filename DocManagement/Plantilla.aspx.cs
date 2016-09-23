using System;
using System.Web.UI.WebControls;
using System.Web.UI;


namespace DocManagement
{
    public partial class Plantilla : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
       
            string cliente = Request.QueryString["id"];
            this.lblCliente.Text = cliente;

            Plantillas Marco = new Plantillas();
            
            for (int i = 0; i < Marco.NumeroPlantillas; i++)
            {
                TableCell c = new TableCell();
                TableRow r = new TableRow();
                HyperLink hyp = new HyperLink();
                hyp.Text = Marco.listado[i].NombrePlantilla;
                hyp.NavigateUrl = Marco.listado[i].UbicacionPlantilla;
                c.Controls.Add(hyp);
                r.Cells.Add(c);
             
                Table1.Rows.Add(r);
            }


        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            string strOrigFile = Server.MapPath("files/originaldoc/CartaPremio.docx");
            string strCopiesDir = Server.MapPath("files/copies");
            string strOutputDir = Server.MapPath("files/output/output.docx");
            DocMerger objMerger = new DocMerger();
            objMerger.Merge(strOrigFile, strCopiesDir, strOutputDir);

        }






    }
}