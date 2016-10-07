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
 
    

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
        public void Guardar()
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

            this.txtParrafo1.Text = "El Pico de Orizaba es una montaña prominente que se localiza en los estados de Puebla y Tlaxcala, México. Ésta se localiza sobre un basamento volcánico de enormes dimensiones y tiene más de un millón y medio de años de antigüedad.";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (this.txtIdCliente.Text == "")
            {
                return;

            }
            this.dataGridView1.DataSource = AccesoDatos.RegresaTablaSql("Select indice as id, fecha, poliza, parrafo1, parrafo2, parrafo3 from CartasClientes where cliente = " + this.txtIdCliente.Text ); ;

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string identificador = "";
            int row = e.RowIndex;
            identificador = dataGridView1.Rows[row].Cells[0].Value.ToString();
            string url = "http://localhost:53816/PlantillaHistorico.aspx?id=" + identificador;

            this.webBrowser1.Navigate(url);
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            if (this.txtPoliza.Text == "")
            { MessageBox.Show("Debe llenar el dato de la poliza."); return; }

            if (this.txtCodigoCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }


            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta) values('" + txtCodigoCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Carta envío ramo 9')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel1.Text = "Carta envío ramo 9";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
            this.webBrowser1.Navigate(url);
            MessageBox.Show("Estamos descargando su documento");
        }

        private void txtCodigoCliente_TextChanged(object sender, EventArgs e)
        {

        }
    }
}








