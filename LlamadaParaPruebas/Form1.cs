using System;
using System.Windows.Forms;
using System.Collections;
using System.Data.SqlClient;
using System.Data;

namespace LlamadaParaPruebas
{
    public partial class Form1 : Form
    {
        bool bandera = false;
        bool historico = false;
        bool llenado = false;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PresentarLinks();


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

            this.dataGridView1.DataSource = AccesoDatos.RegresaTablaSql("Select indice as id,tipo_carta, fecha, poliza, parrafo1, parrafo2, parrafo3 from CartasClientes where cliente = " + this.txtIdCliente.Text); ;
            dataGridView1.AllowUserToAddRows = false;

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string identificador = "";
            int row = e.RowIndex;
            identificador = dataGridView1.Rows[row].Cells[0].Value.ToString();
            historico = true;
            this.lblId.Text = identificador;
            string carta = dataGridView1.Rows[row].Cells[1].Value.ToString(); ;
            if (carta == "")
                return;
            PresentarLinks();
            EsconderLinks(carta);
       


        }
        private void EsconderLinks(string carta)
        {
            if (carta != "Carta envío ramo 9")
            {
                linkLabel1.Visible = false;
            }
            if (carta != "Carta envio ramo 123")
            {
                linkLabel2.Visible = false;
            }
            if (carta != "Carta envio todos excepto 9 y 123")
            {
                linkLabel3.Visible = false;
            }
            if (carta != "Envio endosos ramo 9")
            {
                linkLabel4.Visible = false;
            }
            if (carta != "Envio endosos ramo 123")
            {
                linkLabel5.Visible = false;
            }
            if (carta != "Envio Endosos todos excepto 9 y 123")
            {
                linkLabel6.Visible = false;
            }
            //*************************************************
            if (carta != "Envio de Documentos Incendio")
            {
                linkLabel7.Visible = false;
            }
            if (carta != "Envio de Documentos Equipo Electronico")
            {
                linkLabel8.Visible = false;
            }
            if (carta != "Envio de Documentos en blanco")
            {
                linkLabel9.Visible = false;
            }
            if (carta != "Envio de Documentos Autos")
            {
                linkLabel10.Visible = false;
            }
            if (carta != "Envío de Cartapacio")
            {
                linkLabel11.Visible = false;
            }
            if (carta != "Aviso de Renovación con Condiciones Autos")
            {
                linkLabel12.Visible = false;
            }
            if (carta != "Envios Varios en Blanco")
            {
                linkLabel13.Visible = false;
            }
            if (carta != "Aviso de Renovación con Condiciones Incendio")
            {
                linkLabel14.Visible = false;
            }
            if (carta != "Envío de Renovación Autos y Daños")
            {
                linkLabel15.Visible = false;
            }

            //***********************last ones

            if (carta != "Envío de Endoso Autos y Daños")
            {
                linkLabel16.Visible = false;
            }
            if (carta != "Envío de Renovación  Nueva Autos  Daños")
            {
                linkLabel17.Visible = false;
            }
            if (carta != "Envío de Documentos")
            {
                linkLabel18.Visible = false;
            }


            if (carta != "Envío de liquidación de reclamo")
            {
                linkLabel19.Visible = false;
            }

            if (carta != "Envío de Cheque y Liquidación de reclamo")
            {
                linkLabel20.Visible = false;
            }

            if (carta != "Envío de Facturación")
            {
                linkLabel21.Visible = false;
            }

            if (carta != "Envío de Endoso Vida y Gastos Médicos")
            {
                linkLabel22.Visible = false;
            }

            if (carta != "Envío de Renovación Vida y Gastos Médicos")
            {
                linkLabel23.Visible = false;
            }

 



 


        }
        private void PresentarLinks()
        {
 
                linkLabel1.Visible = true;
 
                linkLabel2.Visible = true;
 
                linkLabel3.Visible = true;
 
                linkLabel4.Visible = true;
 
                linkLabel5.Visible = true;
 
                linkLabel6.Visible = true;

                linkLabel7.Visible = true;

                linkLabel8.Visible = true;

                linkLabel9.Visible = true;

                linkLabel10.Visible = true;

                linkLabel11.Visible = true;

                linkLabel12.Visible = true;

                linkLabel13.Visible = true;

            linkLabel14.Visible = true;

            linkLabel15.Visible = true;

            linkLabel16.Visible = true;

            linkLabel17.Visible = true;

            linkLabel18.Visible = true;

            linkLabel19.Visible = true;

            linkLabel20.Visible = true;

            linkLabel21.Visible = true;

            linkLabel22.Visible = true;

            linkLabel23.Visible = true;
 

        }
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }

            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel1.Text = "Carta envío ramo 9";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                this.lblId.Text = ".";
                return;
            }

            if (this.lblId.Text == "Historico")
            {  return; }

      

            if (this.txtIdCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = "+ this.txtIdCliente.Text +" and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

            if (verificador == "0")
            {
                MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                return;
            }
            else {
                this.lblMsg.Text = ".";
            }

            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Carta envío ramo 9', '"+ this.txtTitulo.Text +"', '"+ this.txtNombreCompleto.Text +"','"+ this.txtDireccion.Text +"' ,'"+ this.txtApartado.Text +"' ,'"+ this.txtBien.Text +"' ,'"+ this.txtAseguradora.Text +"','"+ this.txtVence.Text +"', '"+ this.txtEndoso.Text +"' ,'"+ this.txtRequerimientos.Text +"','"+ this.txtAgente.Text +"','"+ this.txtCorreo.Text +"')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + identificador + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel1.Text = "Carta envío ramo 9";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;

            this.webBrowser1.Navigate(url);
           
            MessageBox.Show("Estamos descargando su documento");

        }

        private void txtCodigoCliente_TextChanged(object sender, EventArgs e)
        {

        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }

            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel2.Text = "Carta envio ramo 123";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                return;
            }

 

            if (this.txtIdCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = " + this.txtIdCliente.Text + " and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

            if (verificador == "0")
            {
                MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                return;
            }
            else
            {
                this.lblMsg.Text = ".";
            }


            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Carta envio ramo 123', '" + this.txtTitulo.Text +"', '"+ this.txtNombreCompleto.Text +"','"+ this.txtDireccion.Text +"' ,'"+ this.txtApartado.Text +"' ,'"+ this.txtBien.Text +"' ,'"+ this.txtAseguradora.Text +"','"+ this.txtVence.Text +"', '"+ this.txtEndoso.Text +"' ,'"+ this.txtRequerimientos.Text +"','"+ this.txtAgente.Text +"','"+ this.txtCorreo.Text +"')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel2.Text = "Carta envio ramo 123";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
            this.webBrowser1.Navigate(url);
            MessageBox.Show("Estamos descargando su documento");
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }

            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel3.Text = "Carta envio todos excepto 9 y 123";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                return;
            }

  

            if (this.txtIdCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = " + this.txtIdCliente.Text + " and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

            if (verificador == "0")
            {
                MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                return;
            }
            else
            {
                this.lblMsg.Text = ".";
            }

            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Carta envio todos excepto 9 y 123', '" + this.txtTitulo.Text + "', '" + this.txtNombreCompleto.Text + "','" + this.txtDireccion.Text + "' ,'" + this.txtApartado.Text + "' ,'" + this.txtBien.Text + "' ,'" + this.txtAseguradora.Text + "','" + this.txtVence.Text + "', '" + this.txtEndoso.Text + "' ,'" + this.txtRequerimientos.Text + "','" + this.txtAgente.Text + "','" + this.txtCorreo.Text + "')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel3.Text = "Carta envio todos excepto 9 y 123";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
            this.webBrowser1.Navigate(url);
            MessageBox.Show("Estamos descargando su documento");
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }

            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel4.Text = "Envio endosos ramo 9";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                return;
            }

 

            if (this.txtIdCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = " + this.txtIdCliente.Text + " and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

            if (verificador == "0")
            {
                MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                return;
            }
            else
            {
                this.lblMsg.Text = ".";
            }


            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Envio endosos ramo 9', '" + this.txtTitulo.Text + "', '" + this.txtNombreCompleto.Text + "','" + this.txtDireccion.Text + "' ,'" + this.txtApartado.Text + "' ,'" + this.txtBien.Text + "' ,'" + this.txtAseguradora.Text + "','" + this.txtVence.Text + "', '" + this.txtEndoso.Text + "' ,'" + this.txtRequerimientos.Text + "','" + this.txtAgente.Text + "','" + this.txtCorreo.Text + "')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel4.Text = "Envio endosos ramo 9";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
            this.webBrowser1.Navigate(url);
            MessageBox.Show("Estamos descargando su documento");
        }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }

            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel5.Text = "Envio endosos ramo 123";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                return;
            }
 

            if (this.txtIdCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = " + this.txtIdCliente.Text + " and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

            if (verificador == "0")
            {
                MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                return;
            }
            else
            {
                this.lblMsg.Text = ".";
            }



            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Envio endosos ramo 123', '" + this.txtTitulo.Text + "', '" + this.txtNombreCompleto.Text + "','" + this.txtDireccion.Text + "' ,'" + this.txtApartado.Text + "' ,'" + this.txtBien.Text + "' ,'" + this.txtAseguradora.Text + "','" + this.txtVence.Text + "', '" + this.txtEndoso.Text + "' ,'" + this.txtRequerimientos.Text + "','" + this.txtAgente.Text + "','" + this.txtCorreo.Text + "')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel5.Text = "Envio endosos ramo 123";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
            this.webBrowser1.Navigate(url);
            MessageBox.Show("Estamos descargando su documento");
        }


        private void linkLabel6_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }

            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel6.Text = "Envio Endosos todos excepto 9 y 123";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                return;
                //}

 

                if (this.txtIdCliente.Text == "")
                { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

                string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = " + this.txtIdCliente.Text + " and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

                if (verificador == "0")
                {
                    MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                    this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                    return;
                }
                else
                {
                    this.lblMsg.Text = ".";
                }

                AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Envio Endosos todos excepto 9 y 123', '" + this.txtTitulo.Text + "', '" + this.txtNombreCompleto.Text + "','" + this.txtDireccion.Text + "' ,'" + this.txtApartado.Text + "' ,'" + this.txtBien.Text + "' ,'" + this.txtAseguradora.Text + "','" + this.txtVence.Text + "', '" + this.txtEndoso.Text + "' ,'" + this.txtRequerimientos.Text + "','" + this.txtAgente.Text + "','" + this.txtCorreo.Text + "')");
                string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
                string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
                AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

                this.lblId.Text = identificador;
                linkLabel6.Text = "Envio Endosos todos excepto 9 y 123";
                string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
                this.webBrowser1.Navigate(url);
                MessageBox.Show("Estamos descargando su documento");
            }
        }

        private void linkLabel12_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }

            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel1.Text = "Aviso de Renovación con Condiciones Autos";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                return;
            }


    

            if (this.txtIdCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = " + this.txtIdCliente.Text + " and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

            if (verificador == "0")
            {
                MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                return;
            }
            else
            {
                this.lblMsg.Text = ".";
            }


            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Aviso de Renovación con Condiciones Autos', '" + this.txtTitulo.Text + "', '" + this.txtNombreCompleto.Text + "','" + this.txtDireccion.Text + "' ,'" + this.txtApartado.Text + "' ,'" + this.txtBien.Text + "' ,'" + this.txtAseguradora.Text + "','" + this.txtVence.Text + "', '" + this.txtEndoso.Text + "' ,'" + this.txtRequerimientos.Text + "','" + this.txtAgente.Text + "','" + this.txtCorreo.Text + "')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel1.Text = "Aviso de Renovación con Condiciones Autos";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
            this.webBrowser1.Navigate(url);
            MessageBox.Show("Estamos descargando su documento");
        }

        private void linkLabel7_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }

            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel1.Text = "Envio de Documentos Incendio";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                return;
            }

 

            if (this.txtIdCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = " + this.txtIdCliente.Text + " and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

            if (verificador == "0")
            {
                MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                return;
            }
            else
            {
                this.lblMsg.Text = ".";
            }


            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Envio de Documentos Incendio', '" + this.txtTitulo.Text + "', '" + this.txtNombreCompleto.Text + "','" + this.txtDireccion.Text + "' ,'" + this.txtApartado.Text + "' ,'" + this.txtBien.Text + "' ,'" + this.txtAseguradora.Text + "','" + this.txtVence.Text + "', '" + this.txtEndoso.Text + "' ,'" + this.txtRequerimientos.Text + "','" + this.txtAgente.Text + "','" + this.txtCorreo.Text + "')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel1.Text = "Envio de Documentos Incendio";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
            this.webBrowser1.Navigate(url);
            MessageBox.Show("Estamos descargando su documento");
        }

        private void linkLabel8_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }

            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel1.Text = "Envio de Documentos Equipo Electronico";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                return;
            }

 

            if (this.txtIdCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = " + this.txtIdCliente.Text + " and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

            if (verificador == "0")
            {
                MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                return;
            }
            else
            {
                this.lblMsg.Text = ".";
            }


            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Envio de Documentos Equipo Electronico', '" + this.txtTitulo.Text + "', '" + this.txtNombreCompleto.Text + "','" + this.txtDireccion.Text + "' ,'" + this.txtApartado.Text + "' ,'" + this.txtBien.Text + "' ,'" + this.txtAseguradora.Text + "','" + this.txtVence.Text + "', '" + this.txtEndoso.Text + "' ,'" + this.txtRequerimientos.Text + "','" + this.txtAgente.Text + "','" + this.txtCorreo.Text + "')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel1.Text = "Envio de Documentos Equipo Electronico";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
            this.webBrowser1.Navigate(url);
            MessageBox.Show("Estamos descargando su documento");
        }

        private void linkLabel9_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }

            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel1.Text = "Envio de Documentos en blanco";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                return;
            }

 

            if (this.txtIdCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = " + this.txtIdCliente.Text + " and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

            if (verificador == "0")
            {
                MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                return;
            }
            else
            {
                this.lblMsg.Text = ".";
            }


            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Envio de Documentos en blanco', '" + this.txtTitulo.Text + "', '" + this.txtNombreCompleto.Text + "','" + this.txtDireccion.Text + "' ,'" + this.txtApartado.Text + "' ,'" + this.txtBien.Text + "' ,'" + this.txtAseguradora.Text + "','" + this.txtVence.Text + "', '" + this.txtEndoso.Text + "' ,'" + this.txtRequerimientos.Text + "','" + this.txtAgente.Text + "','" + this.txtCorreo.Text + "')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel1.Text = "Envio de Documentos en blanco";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
            this.webBrowser1.Navigate(url);
            MessageBox.Show("Estamos descargando su documento");
        }

        private void linkLabel10_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }

            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel1.Text = "Envio de Documentos Autos";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                return;
            }

 

            if (this.txtIdCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = " + this.txtIdCliente.Text + " and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

            if (verificador == "0")
            {
                MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                return;
            }
            else
            {
                this.lblMsg.Text = ".";
            }


            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Envio de Documentos Autos', '" + this.txtTitulo.Text + "', '" + this.txtNombreCompleto.Text + "','" + this.txtDireccion.Text + "' ,'" + this.txtApartado.Text + "' ,'" + this.txtBien.Text + "' ,'" + this.txtAseguradora.Text + "','" + this.txtVence.Text + "', '" + this.txtEndoso.Text + "' ,'" + this.txtRequerimientos.Text + "','" + this.txtAgente.Text + "','" + this.txtCorreo.Text + "')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel1.Text = "Envio de Documentos Autos";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
            this.webBrowser1.Navigate(url);
            MessageBox.Show("Estamos descargando su documento");
        }

        private void linkLabel11_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }

            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel1.Text = "Envío de Cartapacio";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                return;
            }


 

            if (this.txtIdCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = " + this.txtIdCliente.Text + " and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

            if (verificador == "0")
            {
                MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                return;
            }
            else
            {
                this.lblMsg.Text = ".";
            }


            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Envío de Cartapacio', '" + this.txtTitulo.Text + "', '" + this.txtNombreCompleto.Text + "','" + this.txtDireccion.Text + "' ,'" + this.txtApartado.Text + "' ,'" + this.txtBien.Text + "' ,'" + this.txtAseguradora.Text + "','" + this.txtVence.Text + "', '" + this.txtEndoso.Text + "' ,'" + this.txtRequerimientos.Text + "','" + this.txtAgente.Text + "','" + this.txtCorreo.Text + "')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel1.Text = "Envío de Cartapacio";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
            this.webBrowser1.Navigate(url);
            MessageBox.Show("Estamos descargando su documento");
        }

        private void linkLabel13_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }

            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel1.Text = "Envios Varios en Blanco";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                return;
            }


 

            if (this.txtIdCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = " + this.txtIdCliente.Text + " and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

            if (verificador == "0")
            {
                MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                return;
            }
            else
            {
                this.lblMsg.Text = ".";
            }

            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Envios Varios en Blanco', '" + this.txtTitulo.Text + "', '" + this.txtNombreCompleto.Text + "','" + this.txtDireccion.Text + "' ,'" + this.txtApartado.Text + "' ,'" + this.txtBien.Text + "' ,'" + this.txtAseguradora.Text + "','" + this.txtVence.Text + "', '" + this.txtEndoso.Text + "' ,'" + this.txtRequerimientos.Text + "','" + this.txtAgente.Text + "','" + this.txtCorreo.Text + "')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel1.Text = "Envios Varios en Blanco";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
            this.webBrowser1.Navigate(url);
            MessageBox.Show("Estamos descargando su documento");


        }

        private void txtPoliza_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtIdCliente_TextChanged(object sender, EventArgs e)
        {
            if (llenado == false)
            {
                this.txtPoliza.Text = ""; this.lblId.Text = ".";
            }
        }

        private void linkLabel14_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }

            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel1.Text = "Aviso de Renovación con Condiciones Incendio";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                return;
            }

 

            if (this.txtIdCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = " + this.txtIdCliente.Text + " and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

            if (verificador == "0")
            {
                MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                return;
            }
            else
            {
                this.lblMsg.Text = ".";
            }

            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Aviso de Renovación con Condiciones Incendio', '" + this.txtTitulo.Text + "', '" + this.txtNombreCompleto.Text + "','" + this.txtDireccion.Text + "' ,'" + this.txtApartado.Text + "' ,'" + this.txtBien.Text + "' ,'" + this.txtAseguradora.Text + "','" + this.txtVence.Text + "', '" + this.txtEndoso.Text + "' ,'" + this.txtRequerimientos.Text + "','" + this.txtAgente.Text + "','" + this.txtCorreo.Text + "')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel1.Text = "Aviso de Renovación con Condiciones Incendio";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
            this.webBrowser1.Navigate(url);
            MessageBox.Show("Estamos descargando su documento");
        }

        private void linkLabel15_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }

            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel1.Text = "Envío de Renovación Autos y Daños";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                return;
            }

 

            if (this.txtIdCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = " + this.txtIdCliente.Text + " and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

            if (verificador == "0")
            {
                MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                return;
            }
            else
            {
                this.lblMsg.Text = ".";
            }

            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Envío de Renovación Autos y Daños', '" + this.txtTitulo.Text + "', '" + this.txtNombreCompleto.Text + "','" + this.txtDireccion.Text + "' ,'" + this.txtApartado.Text + "' ,'" + this.txtBien.Text + "' ,'" + this.txtAseguradora.Text + "','" + this.txtVence.Text + "', '" + this.txtEndoso.Text + "' ,'" + this.txtRequerimientos.Text + "','" + this.txtAgente.Text + "','" + this.txtCorreo.Text + "')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel1.Text = "Envío de Renovación Autos y Daños";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
            this.webBrowser1.Navigate(url);
            MessageBox.Show("Estamos descargando su documento");
        }

        private void linkLabel16_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }

            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel1.Text = "Envío de Endoso Autos y Daños";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                return;
            }

 

            if (this.txtIdCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = " + this.txtIdCliente.Text + " and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

            if (verificador == "0")
            {
                MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                return;
            }
            else
            {
                this.lblMsg.Text = ".";
            }

            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Envío de Endoso Autos y Daños', '" + this.txtTitulo.Text + "', '" + this.txtNombreCompleto.Text + "','" + this.txtDireccion.Text + "' ,'" + this.txtApartado.Text + "' ,'" + this.txtBien.Text + "' ,'" + this.txtAseguradora.Text + "','" + this.txtVence.Text + "', '" + this.txtEndoso.Text + "' ,'" + this.txtRequerimientos.Text + "','" + this.txtAgente.Text + "','" + this.txtCorreo.Text + "')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel1.Text = "Envío de Endoso Autos y Daños";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
            this.webBrowser1.Navigate(url);
            MessageBox.Show("Estamos descargando su documento");

        }

        private void linkLabel17_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }

            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel1.Text = "Envío de Renovación Vida y Gastos Médicos";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                return;
            }


 

            if (this.txtIdCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = " + this.txtIdCliente.Text + " and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

            if (verificador == "0")
            {
                MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                return;
            }
            else
            {
                this.lblMsg.Text = ".";
            }

            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Envío de Renovación Vida y Gastos Médicos', '" + this.txtTitulo.Text + "', '" + this.txtNombreCompleto.Text + "','" + this.txtDireccion.Text + "' ,'" + this.txtApartado.Text + "' ,'" + this.txtBien.Text + "' ,'" + this.txtAseguradora.Text + "','" + this.txtVence.Text + "', '" + this.txtEndoso.Text + "' ,'" + this.txtRequerimientos.Text + "','" + this.txtAgente.Text + "','" + this.txtCorreo.Text + "')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel1.Text = "Envío de Renovación Vida y Gastos Médicos";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
            this.webBrowser1.Navigate(url);
            MessageBox.Show("Estamos descargando su documento");

        }

        private void linkLabel18_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }

            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel1.Text = "Envío de Documentos";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                return;
            }


 

            if (this.txtIdCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = " + this.txtIdCliente.Text + " and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

            if (verificador == "0")
            {
                MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                return;
            }
            else
            {
                this.lblMsg.Text = ".";
            }

            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Envío de Documentos', '" + this.txtTitulo.Text + "', '" + this.txtNombreCompleto.Text + "','" + this.txtDireccion.Text + "' ,'" + this.txtApartado.Text + "' ,'" + this.txtBien.Text + "' ,'" + this.txtAseguradora.Text + "','" + this.txtVence.Text + "', '" + this.txtEndoso.Text + "' ,'" + this.txtRequerimientos.Text + "','" + this.txtAgente.Text + "','" + this.txtCorreo.Text + "')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel1.Text = "Envío de Documentos";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
            this.webBrowser1.Navigate(url);
            MessageBox.Show("Estamos descargando su documento");

        }

        private void linkLabel19_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }


            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel1.Text = "Envío de liquidación de reclamo";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                return;
            }

 

            if (this.txtIdCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = " + this.txtIdCliente.Text + " and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

            if (verificador == "0")
            {
                MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                return;
            }
            else
            {
                this.lblMsg.Text = ".";
            }

            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Envío de liquidación de reclamo', '" + this.txtTitulo.Text + "', '" + this.txtNombreCompleto.Text + "','" + this.txtDireccion.Text + "' ,'" + this.txtApartado.Text + "' ,'" + this.txtBien.Text + "' ,'" + this.txtAseguradora.Text + "','" + this.txtVence.Text + "', '" + this.txtEndoso.Text + "' ,'" + this.txtRequerimientos.Text + "','" + this.txtAgente.Text + "','" + this.txtCorreo.Text + "')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel1.Text = "Envío de liquidación de reclamo";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
            this.webBrowser1.Navigate(url);
            MessageBox.Show("Estamos descargando su documento");
        }

        private void linkLabel20_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }

            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel1.Text = "Envío de Cheque y Liquidación de reclamo";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                return;
            }


 

            if (this.txtIdCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = " + this.txtIdCliente.Text + " and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

            if (verificador == "0")
            {
                MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                return;
            }
            else
            {
                this.lblMsg.Text = ".";
            }

            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Envío de Cheque y Liquidación de reclamo', '" + this.txtTitulo.Text + "', '" + this.txtNombreCompleto.Text + "','" + this.txtDireccion.Text + "' ,'" + this.txtApartado.Text + "' ,'" + this.txtBien.Text + "' ,'" + this.txtAseguradora.Text + "','" + this.txtVence.Text + "', '" + this.txtEndoso.Text + "' ,'" + this.txtRequerimientos.Text + "','" + this.txtAgente.Text + "','" + this.txtCorreo.Text + "')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel1.Text = "Envío de Cheque y Liquidación de reclamo";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
            this.webBrowser1.Navigate(url);
            MessageBox.Show("Estamos descargando su documento");

        }

        private void linkLabel21_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }

            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel1.Text = "Envío de Facturación";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                return;
            }


 

            if (this.txtIdCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = " + this.txtIdCliente.Text + " and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

            if (verificador == "0")
            {
                MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                return;
            }
            else
            {
                this.lblMsg.Text = ".";
            }

            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Envío de Facturación', '" + this.txtTitulo.Text + "', '" + this.txtNombreCompleto.Text + "','" + this.txtDireccion.Text + "' ,'" + this.txtApartado.Text + "' ,'" + this.txtBien.Text + "' ,'" + this.txtAseguradora.Text + "','" + this.txtVence.Text + "', '" + this.txtEndoso.Text + "' ,'" + this.txtRequerimientos.Text + "','" + this.txtAgente.Text + "','" + this.txtCorreo.Text + "')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel1.Text = "Envío de Facturación";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
            this.webBrowser1.Navigate(url);
            MessageBox.Show("Estamos descargando su documento");
        }

        private void linkLabel22_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }

            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel1.Text = "Envío de Endoso Vida y Gastos Médicos";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                return;
            }

 

            if (this.txtIdCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = " + this.txtIdCliente.Text + " and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

            if (verificador == "0")
            {
                MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                return;
            }
            else
            {
                this.lblMsg.Text = ".";
            }

            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Envío de Endoso Vida y Gastos Médicos', '" + this.txtTitulo.Text + "', '" + this.txtNombreCompleto.Text + "','" + this.txtDireccion.Text + "' ,'" + this.txtApartado.Text + "' ,'" + this.txtBien.Text + "' ,'" + this.txtAseguradora.Text + "','" + this.txtVence.Text + "', '" + this.txtEndoso.Text + "' ,'" + this.txtRequerimientos.Text + "','" + this.txtAgente.Text + "','" + this.txtCorreo.Text + "')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel1.Text = "Envío de Endoso Vida y Gastos Médicos";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
            this.webBrowser1.Navigate(url);
            MessageBox.Show("Estamos descargando su documento");
        }

        private void limipiar_parrafo()
        {
            this.txtParrafo1.Text = "";
            this.txtParrafo2.Text = "";
            this.txtParrafo3.Text = "";
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            limipiar_parrafo();

            this.txtParrafo1.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta = 'Carta envío ramo 9'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta = 'Carta envío ramo 9'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta = 'Carta envío ramo 9'");

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            limipiar_parrafo();
            this.txtParrafo1.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta = 'Carta envio ramo 123'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta = 'Carta envio ramo 123'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta = 'Carta envio ramo 123'");

        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            limipiar_parrafo();
            this.txtParrafo1.Text  = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta = 'Carta envio todos excepto 9 y 123'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta = 'Carta envio todos excepto 9 y 123'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta = 'Carta envio todos excepto 9 y 123'");

        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            limipiar_parrafo();
            this.txtParrafo1.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta = 'Envio Endosos todos excepto 9 y 123'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta = 'Envio Endosos todos excepto 9 y 123'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta = 'Envio Endosos todos excepto 9 y 123'");
           
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            limipiar_parrafo();
            this.txtParrafo1.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta = 'Envio endosos ramo 123'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta = 'Envio endosos ramo 123'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta = 'Envio endosos ramo 123'");

        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            limipiar_parrafo();
            this.txtParrafo1.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta = 'Envio Endosos todos excepto 9 y 123'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta = 'Envio Endosos todos excepto 9 y 123'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta = 'Envio Endosos todos excepto 9 y 123'");

        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            limipiar_parrafo();
            this.txtParrafo1.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta = 'Envio de Documentos Incendio'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta = 'Envio de Documentos Incendio'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta = 'Envio de Documentos Incendio'");
        
        }

        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            limipiar_parrafo();
            this.txtParrafo1.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta ='Envio de Documentos Equipo Electronico'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta = 'Envio de Documentos Equipo Electronico'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta = 'Envio de Documentos Equipo Electronico'");
        }

        private void radioButton9_CheckedChanged(object sender, EventArgs e)
        {
        
            limipiar_parrafo();
            this.txtParrafo1.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta ='Envio de Documentos en blanco'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta = 'Envio de Documentos en blanco'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta = 'Envio de Documentos en blanco'");


        }

        private void radioButton10_CheckedChanged(object sender, EventArgs e)
        {
      
            limipiar_parrafo();
            this.txtParrafo1.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta ='Envio de Documentos Autos'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta = 'Envio de Documentos Autos'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta = 'Envio de Documentos Autos'");



        }

        private void radioButton11_CheckedChanged(object sender, EventArgs e)
        {
           
            limipiar_parrafo();
            this.txtParrafo1.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta ='Envío de Cartapacio'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta = 'Envío de Cartapacio'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta = 'Envío de Cartapacio'");

        }

        private void radioButton12_CheckedChanged(object sender, EventArgs e)
        {
         
            limipiar_parrafo();
            this.txtParrafo1.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta ='Aviso de Renovación con Condiciones Autos'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta = 'Aviso de Renovación con Condiciones Autos'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta ='Aviso de Renovación con Condiciones Autos'");

        }

        private void radioButton13_CheckedChanged(object sender, EventArgs e)
        {
        
            limipiar_parrafo();
            this.txtParrafo1.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta = 'Aviso de Renovación con Condiciones Incendio'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta =  'Aviso de Renovación con Condiciones Incendio'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta = 'Aviso de Renovación con Condiciones Incendio'");

        }

        private void radioButton14_CheckedChanged(object sender, EventArgs e)
        {
           
            limipiar_parrafo();
            this.txtParrafo1.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta =  'Envío de Renovación Autos y Daños'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta =  'Envío de Renovación Autos y Daños'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta =  'Envío de Renovación Autos y Daños'");
        }

        private void radioButton15_CheckedChanged(object sender, EventArgs e)
        {

            limipiar_parrafo();
            this.txtParrafo1.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta =  'Envío de Renovación Vida y Gastos Médicos'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta =  'Envío de Renovación Vida y Gastos Médicos'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta =  'Envío de Renovación Vida y Gastos Médicos'");

        }

        private void radioButton16_CheckedChanged(object sender, EventArgs e)
        {
 
            limipiar_parrafo();
            this.txtParrafo1.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta =  'Envío de Endoso Autos y Daños'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta =  'Envío de Endoso Autos y Daños'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta =  'Envío de Endoso Autos y Daños'");


        }

        private void radioButton17_CheckedChanged(object sender, EventArgs e)
        {
            limipiar_parrafo();
            this.txtParrafo1.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta =  'Envío de Renovación  Nueva Autos  Daños'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta =  'Envío de Renovación  Nueva Autos  Daños'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta =  'Envío de Renovación  Nueva Autos  Daños'");


        }

        private void radioButton18_CheckedChanged(object sender, EventArgs e)
        {
 
            limipiar_parrafo();
            this.txtParrafo1.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta = 'Envío de Documentos'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta =  'Envío de Documentos'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta =  'Envío de Documentos'");
        }

        private void radioButton19_CheckedChanged(object sender, EventArgs e)
        {
            limipiar_parrafo();
            this.txtParrafo1.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta = 'Envío de liquidación de reclamo'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta =  'Envío de liquidación de reclamo'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta =  'Envío de liquidación de reclamo'");

        }

        private void radioButton20_CheckedChanged(object sender, EventArgs e)
        {
            limipiar_parrafo();
            this.txtParrafo1.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta = 'Envío de Cheque y Liquidación de reclamo'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta = 'Envío de Cheque y Liquidación de reclamo'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta = 'Envío de Cheque y Liquidación de reclamo'");


        }

        private void radioButton21_CheckedChanged(object sender, EventArgs e)
        {

            limipiar_parrafo();
            this.txtParrafo1.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta = 'Envío de Facturación'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta = 'Envío de Facturación'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta = 'Envío de Facturación'");

        }

        private void radioButton22_CheckedChanged(object sender, EventArgs e)
        {

            limipiar_parrafo();
            this.txtParrafo1.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta = 'Envío de Endoso Vida y Gastos Médicos'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta = 'Envío de Endoso Vida y Gastos Médicos'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta = 'Envío de Endoso Vida y Gastos Médicos'");

        }

        private void radioButton23_CheckedChanged(object sender, EventArgs e)
        {
         
            limipiar_parrafo();
            this.txtParrafo1.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault1 from CartasClientes_Default where NombreCarta = 'Envios Varios en Blanco'");
            this.txtParrafo2.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault2 from CartasClientes_Default where NombreCarta ='Envios Varios en Blanco'");
            this.txtParrafo3.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select ParrafoDefault3 from CartasClientes_Default where NombreCarta ='Envios Varios en Blanco'");



        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Left = 0;
            this.Top = 0;
        }

        private void linkLabel23_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.txtNombreCompleto.Text == "")
            {
                MessageBox.Show("Porfavor seleccione información para la portada de la carta.");
                return;
            }


            if (historico == true)
            {
                string id_historico = this.lblId.Text;

                linkLabel1.Text = "Envío de Renovación Vida y Gastos Médicos";
                string navegar = "http://localhost:53816/Plantilla1.aspx?id=" + id_historico;
                this.webBrowser1.Navigate(navegar);
                MessageBox.Show("Estamos descargando su documento");
                historico = false;
                return;
            }


 

            if (this.txtIdCliente.Text == "")
            { MessageBox.Show("Debe llenar el codigo del cliente."); return; }

            string verificador = AccesoDatos.RegresaCadena_1_ResultadoSql("select count(*) from poliza where  cliente = " + this.txtIdCliente.Text + " and  poliza = '" + this.txtPoliza.Text + "' and tipo = 'poliza' and status != 'Cancelada'");

            if (verificador == "0")
            {
                MessageBox.Show("No se encontro este numero de poliza o se encuentra cancelada, verifique que el cliente ingresado es el que corresponde a la poliza.");
                this.lblMsg.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("Select top 1 status from poliza where poliza  = '" + this.txtPoliza.Text + "'  and ISNULL(status,'') != '' ");
                return;
            }
            else
            {
                this.lblMsg.Text = ".";
            }

            AccesoDatos.EjecutaQuerySql("insert into CartasClientes(cliente, fecha, poliza, parrafo1, parrafo2, parrafo3,tipo_carta,titulo, nombrecompleto, direccion, apartado, bien, aseguradora, vence, endoso, requerimiento,firma, correo) values('" + txtIdCliente.Text + "', getdate() , '" + this.txtPoliza.Text + "' ,'" + this.txtParrafo1.Text + "', '" + this.txtParrafo2.Text + "', '" + this.txtParrafo3.Text + "','Envío de Renovación Vida y Gastos Médicos', '" + this.txtTitulo.Text + "', '" + this.txtNombreCompleto.Text + "','" + this.txtDireccion.Text + "' ,'" + this.txtApartado.Text + "' ,'" + this.txtBien.Text + "' ,'" + this.txtAseguradora.Text + "','" + this.txtVence.Text + "', '" + this.txtEndoso.Text + "' ,'" + this.txtRequerimientos.Text + "','" + this.txtAgente.Text + "','" + this.txtCorreo.Text + "')");
            string identificador = AccesoDatos.RegresaCadena_1_ResultadoSql("Select max(indice) from CartasClientes;");
            string codigo = AccesoDatos.RegresaCadena_1_ResultadoSql("DECLARE @myid uniqueidentifier = NEWID(); SELECT Replace(left(CONVERT(char(255), @myid),10),'-','8') + '-' + '" + identificador + "' AS 'char'; ");
            AccesoDatos.EjecutaQuerySql("Update CartasClientes Set codigo = '" + codigo + "' where indice =" + identificador);

            this.lblId.Text = identificador;
            linkLabel1.Text = "Envío de Renovación Vida y Gastos Médicos";
            string url = "http://localhost:53816/Plantilla1.aspx?id=" + identificador;
            this.webBrowser1.Navigate(url);
            MessageBox.Show("Estamos descargando su documento");
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (this.comboBox1.Text.ToString() == "")
            {
                MessageBox.Show("Debe llenar la secuencia de renovación");
                return;
            }

                if (this.txtPoliza.Text != "")
            {
                llenado = true;
                this.txtIdCliente.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("select top 1 cliente from poliza where  poliza = '" + this.txtPoliza.Text.ToString().Trim().Replace("'", "") + "' and tipo = 'poliza' ");
                this.lblId.Text = ".";
                llenado = false;
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
             "    where p.poliza = '" + this.txtPoliza.Text.ToString().Trim().Replace("'", "") + "' and secren =  '"+ this.comboBox1.Text.ToString()  +"' and p.tipo = 'poliza' and p.status != 'Cancelada' and secren = (select MAX(secren) from poliza where poliza ='" + this.txtPoliza.Text.ToString().Trim().Replace("'", "") + "' and tipo = 'poliza' and status != 'Cancelada' )");

                foreach (System.Data.DataRow rw in content.Rows)
                {
                
                    this.txtTitulo.Text = rw["Titulo"].ToString() ;
                    this.txtNombreCompleto.Text = rw["NombreCompleto"].ToString();
                    this.txtDireccion.Text = rw["direccion"].ToString();
                    this.txtApartado.Text = rw["apartado"].ToString();
                    this.txtBien.Text = rw["descr"].ToString();
                    this.txtAseguradora.Text = rw["nombre"].ToString();
                    this.txtVence.Text = rw["fecha_vencimiento_contrato"].ToString().Substring(0,10);
                    this.txtEndoso.Text = rw["endoso"].ToString();
                    this.txtAgente.Text = rw["gst_nombre"].ToString();
                    this.txtCorreo.Text = rw["gst_correo"].ToString() ;

                    break;
                }

                System.Data.DataTable request = new System.Data.DataTable();
                try
                {
                    if (this.panel1.Visible == true)
                    {
                        request = AccesoDatos.RegresaTablaSql("select numero_oficina from cmorosidad  where poliza = '" + this.txtPoliza.Text.Trim().Replace("'", "") + "' and secren = " + this.comboBox1.Text + "  and ramo = " + this.comboBox3.SelectedValue.ToString());
                    }

                    if (this.panel2.Visible == true)
                    {
                        request = AccesoDatos.RegresaTablaSql("select numero_oficina from cmorosidad  where poliza = '" + this.comboBox2.Text.Trim().Replace("'", "") + "' and secren = " + this.comboBox1.Text + "  and ramo = " + this.comboBox3.SelectedValue.ToString());

                    }
                }
                catch (Exception)
                {

                    request = AccesoDatos.RegresaTablaSql("select numero_oficina from cmorosidad  where poliza = '" + this.comboBox2.Text.Trim().Replace("'", "") + "' and secren = " + this.comboBox1.Text );
                }

                this.txtRequerimientos.Text = "";

                foreach (System.Data.DataRow rw in request.Rows)
                {
                    if (this.txtRequerimientos.Text == "")
                    {
                        this.txtRequerimientos.Text += rw["numero_oficina"].ToString();
                    }
                    else {
                        this.txtRequerimientos.Text += "," + rw["numero_oficina"].ToString();
                    }

                    
                }

                if (content.Rows.Count == 0)
                {
                    limpiar();
                }



            }


        }


        public void limpiar()
        {

            this.txtTitulo.Text = "";
            this.txtNombreCompleto.Text = "";
            this.txtDireccion.Text = "";
            this.txtApartado.Text = "";
            this.txtBien.Text = "";
            this.txtAseguradora.Text = "";
            this.txtVence.Text = "";
            this.txtEndoso.Text = "";
            this.txtAgente.Text = "";
            this.txtCorreo.Text = "";
            this.txtRequerimientos.Text = "";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.panel1.Visible = false;
            this.panel2.Visible = true;
            if (this.txtIdCliente.Text == "")
            {
                MessageBox.Show("Debe indicar que codigo de cliente quiere visualizar.");
                return;
            }


            this.comboBox2.DataSource = AccesoDatos.RegresaTablaSql("select poliza from poliza where cliente = '"+ this.txtIdCliente.Text +"' group by poliza");
            this.comboBox2.DisplayMember = "poliza";
            this.comboBox2.Refresh();
            this.checkBox1.Checked = false;

        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (this.comboBox2.Text.ToString() == "")
            {
                MessageBox.Show("Seleccione una poliza del cliente o es posible que el cliente no tiene polizas registradas.");
                return;
            }

            if (this.comboBox1.Text.ToString() == "")
            {
                MessageBox.Show("Debe llenar la secuencia de renovación");
                return;
            }

            if (this.comboBox2.Text.ToString()  != "")
            {
                llenado = true;
                this.txtIdCliente.Text = AccesoDatos.RegresaCadena_1_ResultadoSql("select top 1 cliente from poliza where  poliza = '" + this.comboBox2.Text.ToString().Trim().Replace("'", "") + "' and tipo = 'poliza' ");
                this.lblId.Text = ".";
                llenado = false;
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
             "    where p.poliza = '" + this.comboBox2.Text.ToString().Trim().Replace("'", "") + "' and secren =  '" + this.comboBox1.Text.ToString().Trim() + "' and p.tipo = 'poliza' and p.status != 'Cancelada' and secren = (select MAX(secren) from poliza where poliza ='" + this.comboBox2.Text.ToString().Trim().Replace("'", "") + "' and tipo = 'poliza' and status != 'Cancelada' )");

                foreach (System.Data.DataRow rw in content.Rows)
                {

                    this.txtTitulo.Text = rw["Titulo"].ToString();
                    this.txtNombreCompleto.Text = rw["NombreCompleto"].ToString();
                    this.txtDireccion.Text = rw["direccion"].ToString();
                    this.txtApartado.Text = rw["apartado"].ToString();
                    this.txtBien.Text = rw["descr"].ToString();
                    this.txtAseguradora.Text = rw["nombre"].ToString();
                    this.txtVence.Text = rw["fecha_vencimiento_contrato"].ToString().Substring(0, 10);
                    this.txtEndoso.Text = rw["endoso"].ToString();
                    this.txtAgente.Text = rw["gst_nombre"].ToString();
                    this.txtCorreo.Text = rw["gst_correo"].ToString();

                    break;
                }
                System.Data.DataTable request = new System.Data.DataTable();
                if (this.panel1.Visible == true)
                {
                    request = AccesoDatos.RegresaTablaSql("select numero_oficina from cmorosidad  where poliza = '" + this.txtPoliza.Text.Trim().Replace("'", "") + "' and secren = " + this.comboBox1.Text + "  and ramo = " + this.comboBox3.SelectedValue.ToString());
                }

                if (this.panel2.Visible == true)
                {
                    request = AccesoDatos.RegresaTablaSql("select numero_oficina from cmorosidad  where poliza = '" + this.comboBox2.Text.Trim().Replace("'", "") + "' and secren = " + this.comboBox1.Text + "  and ramo = " + this.comboBox3.SelectedValue.ToString());

                }
                this.txtRequerimientos.Text = "";

                foreach (System.Data.DataRow rw in request.Rows)
                {
                    if (this.txtRequerimientos.Text == "")
                    {
                        this.txtRequerimientos.Text += rw["numero_oficina"].ToString();
                    }
                    else
                    {
                        this.txtRequerimientos.Text += "," + rw["numero_oficina"].ToString();
                    }


                }

                if (content.Rows.Count == 0)
                {
                    limpiar();
                }


            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.panel1.Visible == true)
            {
                this.comboBox3.DataSource = AccesoDatos.RegresaTablaSql("select descr, ramo from ramos where ramo in (select ramo from cmorosidad where poliza = '" + this.txtPoliza.Text.Trim().Replace("'", "") + "' and secren = " + this.comboBox1.Text.Trim() +" )");
                this.comboBox3.DisplayMember = "descr";
                this.comboBox3.ValueMember = "ramo";
                this.comboBox3.Refresh();
            }

            if (this.panel2.Visible == true)
            {
                this.comboBox3.DataSource = AccesoDatos.RegresaTablaSql("select descr, ramo from ramos where ramo in (select ramo from cmorosidad where poliza = '" + this.comboBox2.Text.Trim().Replace("'", "") + "' and secren = " + this.comboBox1.Text.Trim() + " )");
                this.comboBox3.DisplayMember = "descr";
                this.comboBox3.ValueMember = "ramo";
                this.comboBox3.Refresh();
            }


        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                this.panel1.Visible = true;
                this.panel2.Visible = false;
            }
            else {
                this.panel1.Visible = false;
                this.panel2.Visible = true;
            }
        }
    }
}








