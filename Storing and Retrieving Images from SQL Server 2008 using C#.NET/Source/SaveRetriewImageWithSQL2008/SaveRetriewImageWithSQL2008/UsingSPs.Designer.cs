namespace SaveRetriewImageWithSQL2008
{
    partial class UsingSPs
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
            this.btnDisplayImage = new System.Windows.Forms.Button();
            this.cmbImageID = new System.Windows.Forms.ComboBox();
            this.btnLoadAndSave = new System.Windows.Forms.Button();
            this.picImage = new System.Windows.Forms.PictureBox();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.grbPicBox = new System.Windows.Forms.GroupBox();
            ((System.ComponentModel.ISupportInitialize)(this.picImage)).BeginInit();
            this.grbPicBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnDisplayImage
            // 
            this.btnDisplayImage.Location = new System.Drawing.Point(125, 49);
            this.btnDisplayImage.Name = "btnDisplayImage";
            this.btnDisplayImage.Size = new System.Drawing.Size(83, 23);
            this.btnDisplayImage.TabIndex = 5;
            this.btnDisplayImage.Text = "Display Image";
            this.btnDisplayImage.UseVisualStyleBackColor = true;
            this.btnDisplayImage.Click += new System.EventHandler(this.btnDisplayImage_Click);
            // 
            // cmbImageID
            // 
            this.cmbImageID.FormattingEnabled = true;
            this.cmbImageID.Location = new System.Drawing.Point(12, 50);
            this.cmbImageID.Name = "cmbImageID";
            this.cmbImageID.Size = new System.Drawing.Size(53, 21);
            this.cmbImageID.TabIndex = 4;
            // 
            // btnLoadAndSave
            // 
            this.btnLoadAndSave.Location = new System.Drawing.Point(12, 12);
            this.btnLoadAndSave.Name = "btnLoadAndSave";
            this.btnLoadAndSave.Size = new System.Drawing.Size(196, 23);
            this.btnLoadAndSave.TabIndex = 3;
            this.btnLoadAndSave.Text = "<<--Load and Save Image-->>";
            this.btnLoadAndSave.UseVisualStyleBackColor = true;
            this.btnLoadAndSave.Click += new System.EventHandler(this.btnLoadAndSave_Click);
            // 
            // picImage
            // 
            this.picImage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.picImage.Location = new System.Drawing.Point(3, 16);
            this.picImage.Name = "picImage";
            this.picImage.Size = new System.Drawing.Size(362, 294);
            this.picImage.TabIndex = 6;
            this.picImage.TabStop = false;
            // 
            // btnRefresh
            // 
            this.btnRefresh.Location = new System.Drawing.Point(71, 48);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(48, 23);
            this.btnRefresh.TabIndex = 5;
            this.btnRefresh.Text = "Refresh";
            this.btnRefresh.UseVisualStyleBackColor = true;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // grbPicBox
            // 
            this.grbPicBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.grbPicBox.Controls.Add(this.picImage);
            this.grbPicBox.Location = new System.Drawing.Point(214, 12);
            this.grbPicBox.Name = "grbPicBox";
            this.grbPicBox.Size = new System.Drawing.Size(368, 313);
            this.grbPicBox.TabIndex = 7;
            this.grbPicBox.TabStop = false;
            this.grbPicBox.Text = "Image Display";
            // 
            // UsingSPs
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(585, 328);
            this.Controls.Add(this.grbPicBox);
            this.Controls.Add(this.btnRefresh);
            this.Controls.Add(this.btnDisplayImage);
            this.Controls.Add(this.cmbImageID);
            this.Controls.Add(this.btnLoadAndSave);
            this.Name = "UsingSPs";
            this.Text = "Storing and Retrieving Images from SQL Server using C#.NET";
            this.Load += new System.EventHandler(this.UsingSPs_Load);
            ((System.ComponentModel.ISupportInitialize)(this.picImage)).EndInit();
            this.grbPicBox.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnDisplayImage;
        private System.Windows.Forms.ComboBox cmbImageID;
        private System.Windows.Forms.Button btnLoadAndSave;
        private System.Windows.Forms.PictureBox picImage;
        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.GroupBox grbPicBox;
    }
}