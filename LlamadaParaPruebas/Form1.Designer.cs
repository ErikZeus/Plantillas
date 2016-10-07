namespace LlamadaParaPruebas
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.txtCodigoCliente = new System.Windows.Forms.TextBox();
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtPoliza = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtParrafo1 = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.button3 = new System.Windows.Forms.Button();
            this.txtIdCliente = new System.Windows.Forms.TextBox();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.panel1 = new System.Windows.Forms.Panel();
            this.linkLabel6 = new System.Windows.Forms.LinkLabel();
            this.linkLabel5 = new System.Windows.Forms.LinkLabel();
            this.linkLabel4 = new System.Windows.Forms.LinkLabel();
            this.linkLabel3 = new System.Windows.Forms.LinkLabel();
            this.linkLabel2 = new System.Windows.Forms.LinkLabel();
            this.label6 = new System.Windows.Forms.Label();
            this.lblId = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.label7 = new System.Windows.Forms.Label();
            this.txtParrafo2 = new System.Windows.Forms.TextBox();
            this.txtParrafo3 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtCodigoCliente
            // 
            this.txtCodigoCliente.Location = new System.Drawing.Point(400, 44);
            this.txtCodigoCliente.Name = "txtCodigoCliente";
            this.txtCodigoCliente.Size = new System.Drawing.Size(118, 20);
            this.txtCodigoCliente.TabIndex = 1;
            this.txtCodigoCliente.TextChanged += new System.EventHandler(this.txtCodigoCliente_TextChanged);
            // 
            // webBrowser1
            // 
            this.webBrowser1.Location = new System.Drawing.Point(12, 367);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.Size = new System.Drawing.Size(20, 20);
            this.webBrowser1.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(400, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Codigo Cliente";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(551, 25);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(35, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Poliza";
            // 
            // txtPoliza
            // 
            this.txtPoliza.Location = new System.Drawing.Point(554, 41);
            this.txtPoliza.Name = "txtPoliza";
            this.txtPoliza.Size = new System.Drawing.Size(118, 20);
            this.txtPoliza.TabIndex = 4;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(401, 80);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Parrafo  1";
            this.label4.Click += new System.EventHandler(this.label4_Click);
            // 
            // txtParrafo1
            // 
            this.txtParrafo1.Location = new System.Drawing.Point(403, 106);
            this.txtParrafo1.MaxLength = 230;
            this.txtParrafo1.Multiline = true;
            this.txtParrafo1.Name = "txtParrafo1";
            this.txtParrafo1.Size = new System.Drawing.Size(409, 48);
            this.txtParrafo1.TabIndex = 9;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(719, 167);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(90, 23);
            this.button2.TabIndex = 11;
            this.button2.Text = "Llenar Parrafo:";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 413);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(797, 207);
            this.dataGridView1.TabIndex = 12;
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(527, 384);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(131, 23);
            this.button3.TabIndex = 13;
            this.button3.Text = "Buscar Id Cliente";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // txtIdCliente
            // 
            this.txtIdCliente.Location = new System.Drawing.Point(403, 384);
            this.txtIdCliente.Name = "txtIdCliente";
            this.txtIdCliente.Size = new System.Drawing.Size(118, 20);
            this.txtIdCliente.TabIndex = 14;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.panel1.Controls.Add(this.linkLabel6);
            this.panel1.Controls.Add(this.linkLabel5);
            this.panel1.Controls.Add(this.linkLabel4);
            this.panel1.Controls.Add(this.linkLabel3);
            this.panel1.Controls.Add(this.linkLabel2);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.lblId);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.linkLabel1);
            this.panel1.Location = new System.Drawing.Point(12, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(382, 349);
            this.panel1.TabIndex = 15;
            // 
            // linkLabel6
            // 
            this.linkLabel6.AutoSize = true;
            this.linkLabel6.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabel6.Location = new System.Drawing.Point(30, 185);
            this.linkLabel6.Name = "linkLabel6";
            this.linkLabel6.Size = new System.Drawing.Size(245, 17);
            this.linkLabel6.TabIndex = 8;
            this.linkLabel6.TabStop = true;
            this.linkLabel6.Text = "Envio Endosos todos excepto 9 y 123";
            // 
            // linkLabel5
            // 
            this.linkLabel5.AutoSize = true;
            this.linkLabel5.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabel5.Location = new System.Drawing.Point(30, 154);
            this.linkLabel5.Name = "linkLabel5";
            this.linkLabel5.Size = new System.Drawing.Size(165, 17);
            this.linkLabel5.TabIndex = 7;
            this.linkLabel5.TabStop = true;
            this.linkLabel5.Text = "Envio endosos ramo 123";
            // 
            // linkLabel4
            // 
            this.linkLabel4.AutoSize = true;
            this.linkLabel4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabel4.Location = new System.Drawing.Point(30, 125);
            this.linkLabel4.Name = "linkLabel4";
            this.linkLabel4.Size = new System.Drawing.Size(149, 17);
            this.linkLabel4.TabIndex = 6;
            this.linkLabel4.TabStop = true;
            this.linkLabel4.Text = "Envio endosos ramo 9";
            // 
            // linkLabel3
            // 
            this.linkLabel3.AutoSize = true;
            this.linkLabel3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabel3.Location = new System.Drawing.Point(30, 96);
            this.linkLabel3.Name = "linkLabel3";
            this.linkLabel3.Size = new System.Drawing.Size(223, 17);
            this.linkLabel3.TabIndex = 5;
            this.linkLabel3.TabStop = true;
            this.linkLabel3.Text = "Carta envio todos excepto 9 y 123";
            // 
            // linkLabel2
            // 
            this.linkLabel2.AutoSize = true;
            this.linkLabel2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabel2.Location = new System.Drawing.Point(30, 66);
            this.linkLabel2.Name = "linkLabel2";
            this.linkLabel2.Size = new System.Drawing.Size(144, 17);
            this.linkLabel2.TabIndex = 4;
            this.linkLabel2.TabStop = true;
            this.linkLabel2.Text = "Carta envio ramo 123";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(33, 322);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(246, 17);
            this.label6.TabIndex = 3;
            this.label6.Text = "Generando el id para esta carta no.  :";
            // 
            // lblId
            // 
            this.lblId.AutoSize = true;
            this.lblId.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblId.Location = new System.Drawing.Point(295, 322);
            this.lblId.Name = "lblId";
            this.lblId.Size = new System.Drawing.Size(12, 17);
            this.lblId.TabIndex = 2;
            this.lblId.Text = ".";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(33, 12);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(64, 17);
            this.label5.TabIndex = 1;
            this.label5.Text = "Plantillas";
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabel1.Location = new System.Drawing.Point(30, 38);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(128, 17);
            this.linkLabel1.TabIndex = 0;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "Carta envío ramo 9";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(402, 168);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(53, 13);
            this.label7.TabIndex = 16;
            this.label7.Text = "Parrafo  2";
            // 
            // txtParrafo2
            // 
            this.txtParrafo2.Location = new System.Drawing.Point(400, 196);
            this.txtParrafo2.MaxLength = 230;
            this.txtParrafo2.Multiline = true;
            this.txtParrafo2.Name = "txtParrafo2";
            this.txtParrafo2.Size = new System.Drawing.Size(409, 48);
            this.txtParrafo2.TabIndex = 17;
            // 
            // txtParrafo3
            // 
            this.txtParrafo3.Location = new System.Drawing.Point(400, 294);
            this.txtParrafo3.MaxLength = 230;
            this.txtParrafo3.Multiline = true;
            this.txtParrafo3.Name = "txtParrafo3";
            this.txtParrafo3.Size = new System.Drawing.Size(409, 48);
            this.txtParrafo3.TabIndex = 19;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(402, 266);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 13);
            this.label3.TabIndex = 18;
            this.label3.Text = "Parrafo  3";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(824, 648);
            this.Controls.Add(this.txtParrafo3);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtParrafo2);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.txtIdCliente);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.txtParrafo1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtPoliza);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.webBrowser1);
            this.Controls.Add(this.txtCodigoCliente);
            this.Name = "Form1";
            this.Text = "Plantillas";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox txtCodigoCliente;
        private System.Windows.Forms.WebBrowser webBrowser1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtPoliza;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtParrafo1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.TextBox txtIdCliente;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label lblId;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.LinkLabel linkLabel6;
        private System.Windows.Forms.LinkLabel linkLabel5;
        private System.Windows.Forms.LinkLabel linkLabel4;
        private System.Windows.Forms.LinkLabel linkLabel3;
        private System.Windows.Forms.LinkLabel linkLabel2;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtParrafo2;
        private System.Windows.Forms.TextBox txtParrafo3;
        private System.Windows.Forms.Label label3;
    }
}

