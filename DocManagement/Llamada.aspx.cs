using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace DocManagement
{
    public partial class Llamada : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            this.BarcodeImage.Visible = false;

        }

        protected void btnEncode_Click(object sender, EventArgs e)
        {
            string strImageURL = "GenerateBarcodeImage.aspx?d=" + this.txtData.Text.Trim();
            this.BarcodeImage.ImageUrl = strImageURL;
            this.BarcodeImage.Width = Convert.ToInt32("250");
            this.BarcodeImage.Height = Convert.ToInt32("25");
            this.BarcodeImage.Visible = true;
        }
    }
}