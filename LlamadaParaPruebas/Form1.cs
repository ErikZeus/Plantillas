using System;
using System.Windows.Forms;

namespace LlamadaParaPruebas
{
    public partial class Form1 : Form
    {
        bool bandera = false;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.txtAsistente.Text == "")
            { MessageBox.Show("Debe llenar quien asistio al cliente.");  return; }

            if (this.txtPoliza.Text == "")
            { MessageBox.Show("Debe llenar el dato de la poliza."); return; }

            if (this.txtCodigoCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            if (this.txtCodigoCliente.Text == "")
            { MessageBox.Show("Debe llenar el cuerpo del mensaje principal."); return; }

   
            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, cuerpo, poliza, asistente) values('"+ txtCodigoCliente.Text +"', getdate() , '"+ this.txtCuerpo.Text +"', '"+ this.txtPoliza.Text +"', '"+ this.txtAsistente.Text + "');");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string url = "http://localhost:53816/Plantilla.aspx?id=" + identificador;

      
            this.webBrowser1.Navigate(url);

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
        public void Guardar()
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

            this.txtCuerpo.Text = "El Pico de Orizaba es una montaña prominente que se localiza en los estados de Puebla y Tlaxcala, México. Ésta se localiza sobre un basamento volcánico de enormes dimensiones y tiene más de un millón y medio de años de antigüedad. En el flanco norte del Pico de Orizaba abundan las navajillas de obsidiana, en el sur destaca la presencia de cerámica que marcan una ruta o camino procesional que conduce a cotas más altas hasta la cumbre. En el oriente el sitio OR-13 muestra un xicalli, cerámica abundante de distintas formas y tipos, además de navajillas de obsidiana. " +
             "  " + " Extracto del Libro Mexico Colinas y Relieve.";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (this.txtIdCliente.Text == "")
            {
                return;

            }
            this.dataGridView1.DataSource = AccesoDatos.RegresaTablaSql("Select indice as id, fecha, poliza from CartasClientes where cliente = " + this.txtIdCliente.Text ); ;

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string identificador = "";
            int row = e.RowIndex;
            identificador = dataGridView1.Rows[row].Cells[0].Value.ToString();
            string url = "http://localhost:53816/PlantillaHistorico.aspx?id=" + identificador;

            this.webBrowser1.Navigate(url);
 

        }


    }
}








