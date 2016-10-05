using System;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;


public class ProcesoDeImagenes
{

    public static void GuardaImagen(string path)
    {
        SqlConnection con = new SqlConnection(DBHandler.GetConnectionString()); //connection to the your database
        FileStream FS = new FileStream(path, FileMode.Open, FileAccess.Read); //create a file stream object associate to user selected file 
        byte[] img = new byte[FS.Length]; //create a byte array with size of user select file stream length
        FS.Read(img, 0, Convert.ToInt32(FS.Length));//read user selected file stream in to byte array

        if (con.State == ConnectionState.Closed)//check whether connection to database is close or not
            con.Open();//if connection is close then only open the connection
        SqlCommand cmd = new SqlCommand("SaveImage", con);//create a SQL command object by passing name of the stored procedure and database connection 
        cmd.CommandType = CommandType.StoredProcedure; //set command object command type to stored procedure type
        cmd.Parameters.Add("@img", SqlDbType.Image).Value = img;//add parameter to the command object and set value to that parameter
        cmd.ExecuteNonQuery();//execute command   

    }

    public static Image VisualizarImagen(string _id)
    {
        Image result = null;

        SqlConnection con = new SqlConnection(DBHandler.GetConnectionString());//connection to the your database
            SqlCommand cmd = new SqlCommand("ReadImage", con);//create a SQL command object by passing name of the stored procedure and database connection 
            cmd.CommandType = CommandType.StoredProcedure; //set command object command type to stored procedure type
            cmd.Parameters.Add("@imgId", SqlDbType.Int).Value = Convert.ToInt32(_id);//add parameter to the command object and set value to that parameter
            SqlDataAdapter adp = new SqlDataAdapter(cmd);//create SQL data adapter with command object
            DataTable dt = new DataTable();//create a data table to hold result of the command
            try
            {
                if (con.State == ConnectionState.Closed)//check whether connection to database is close or not
                    con.Open();//if connection is close then only open the connection
                adp.Fill(dt);//data table fill with data by calling the fill method of data adapter object
                if (dt.Rows.Count > 0)//check whether data table contain any row or not
                {
                    MemoryStream ms = new MemoryStream((byte[])dt.Rows[0]["ImageData"]);//create memory stream by passing byte array of the image
                result = Image.FromStream(ms);//set image property of the picture box by creating a image from stream 
                return result;
            }
            }
            catch (Exception ex)//catch if any error occur
            {
            throw;
            }
            finally
            {
                if (con.State == ConnectionState.Open)//check whether connection to database is open or not
                    con.Close();//if connection is open then only close the connection
            }
    
        return result;

    }



}

