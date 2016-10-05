/*
 * Author: Thilina Chandima
 * License : Free to Use Anybody
 * Date of Create : 26/03/2012
 * Time of Create : 09.09 PM
 * File Name : UsingSPs.cs
 * Contact : tcgunarathena@gmail.com
 * 
 * NOTE: Appreciate Any Suggestions or Comments 
 * */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;

namespace SaveRetriewImageWithSQL2008
{
    public partial class UsingSPs : Form
    {
        #region Frm UsingSPs Constructer
        public UsingSPs()
        {
            InitializeComponent();
        } 
        #endregion

        #region Btn Load and Save Click
        private void btnLoadAndSave_Click(object sender, EventArgs e)
        {
            #region Save Image to DB
            SqlConnection con = new SqlConnection(DBHandler.GetConnectionString()); //connection to the your database
            try
            {
                OpenFileDialog fop = new OpenFileDialog(); //create object of open file dialog
                fop.InitialDirectory = @"C:\"; //set Initial directory
                fop.Filter = "[JPG,JPEG]|*.jpg"; //set filter for select only .jpg files
                if (fop.ShowDialog() == DialogResult.OK) //display open file dialog to user and only user select a image enter to if block
                { //@fop.FileName
                    ProcesoDeImagenes.GuardaImagen(@fop.FileName);                     
                    loadImageIDs();//call user defined method to load image IDs to combo box
                    MessageBox.Show("Image Save Successfully!!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);//display save successful message to user
                }
                else
                {
                    MessageBox.Show("Please Select a Image to save!!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);//display message to force select a image 
                }

            }
            catch (Exception ex)//catch if any error occur
            {
                MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);//display error message with exception 
            }
            finally
            {
                if (con.State == ConnectionState.Open)//check whether connection to database is open or not
                    con.Close();//if connection is open then only close the connection
            }
            #endregion
        } 
        #endregion

        #region Btn Display Image Click
        private void btnDisplayImage_Click(object sender, EventArgs e)
        {
            #region Retrieve Image from DB
            if (cmbImageID.SelectedValue != null)//check whether user select image ID or not 
            {
                if (picImage.Image != null)//check whether picture box contain image or not
                    picImage.Image.Dispose();//clear the image of the picture box if there is image

                string _id = cmbImageID.Text;
                picImage.Image = ProcesoDeImagenes.VisualizarImagen(_id);
                picImage.SizeMode = PictureBoxSizeMode.StretchImage;//set size mode property of the picture box to stretch 
                picImage.Refresh();//refresh picture box

            }
            else
            {
                MessageBox.Show("Please Select a Image ID to Display!!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);//display message to force select a image ID
            }
            #endregion
        } 
        #endregion

        #region Btn Refresh Click
        private void btnRefresh_Click(object sender, EventArgs e)
        {
            loadImageIDs();//call user defined method to load image IDs to combo box
        } 
        #endregion

        #region Frm UsingSPs Load
        private void UsingSPs_Load(object sender, EventArgs e)
        {
            loadImageIDs();//call user defined method to load image IDs to combo box
        } 
        #endregion

        #region User Define LoadImageIDs Method
        private void loadImageIDs()
        {
            #region Load Image Ids
            SqlConnection con = new SqlConnection(DBHandler.GetConnectionString());//connection to the your database
            SqlCommand cmd = new SqlCommand("ReadAllImageIDs", con);//create a SQL command object by passing name of the stored procedure and database connection 
            cmd.CommandType = CommandType.StoredProcedure;//set command object command type to stored procedure type
            SqlDataAdapter adp = new SqlDataAdapter(cmd);//create SQL data adapter with command object
            DataTable dt = new DataTable();//create a data table to hold result of the command
            try
            {
                if (con.State == ConnectionState.Closed)//check whether connection to database is close or not
                    con.Open();//if connection is close then only open the connection
                adp.Fill(dt);//data table fill with data by calling the fill method of data adapter object
                if (dt.Rows.Count > 0)//check whether data table contain any row or not
                {
                    cmbImageID.DataSource = dt;//set the data source property of the combo box to result set of the command 
                    cmbImageID.ValueMember = "ImageID";//set the value member property of the combo box
                    cmbImageID.DisplayMember = "ImageID";//set the display member property of the combo box
                    cmbImageID.SelectedIndex = 0;//set the selected index property of the combo box to 0
                }
            }
            catch (Exception ex)//catch if any error occur
            {
                MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);//display error message with exception 
            }
            finally
            {
                if (con.State == ConnectionState.Open)//check whether connection to database is open or not
                    con.Close();//if connection is open then only close the connection
            }
            #endregion
        } 
        #endregion


    }
}
