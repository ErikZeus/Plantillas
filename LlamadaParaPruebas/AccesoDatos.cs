using System;
using System.Data;
using System.Xml;
using System.Xml.Linq;
using System.IO;
using System.Text;

namespace LlamadaParaPruebas
{
    class AccesoDatos
    {
        public static string RegresaCadena_1_ResultadoSql(string sql)
        {

            System.Data.SqlClient.SqlConnection cnn = new System.Data.SqlClient.SqlConnection(CadenaConexion());

            System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(sql, cnn);
            cmd.CommandType = CommandType.Text;
            System.Data.SqlClient.SqlDataAdapter adpt = new System.Data.SqlClient.SqlDataAdapter(cmd);
            System.Data.DataTable content = new System.Data.DataTable();

            cnn.Open();
            adpt.Fill(content);
            cnn.Close();

            string result = "";
            foreach (DataRow rw in content.Rows)
            {
                result = rw[0].ToString();
                break;
            }

            return result;
        }

        public static DataTable RegresaTablaSql(string sql)
        {

            System.Data.SqlClient.SqlConnection cnn = new System.Data.SqlClient.SqlConnection(CadenaConexion());

            System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(sql, cnn);
            cmd.CommandType = CommandType.Text;
            System.Data.SqlClient.SqlDataAdapter adpt = new System.Data.SqlClient.SqlDataAdapter(cmd);
            System.Data.DataTable content = new System.Data.DataTable();

            cnn.Open();
            adpt.Fill(content);
            cnn.Close();

            return content;
        }
        public static string CadenaConexion()
        {
            string conStr = "";
            string path = AppDomain.CurrentDomain.BaseDirectory;
            path = path.Replace(@"\bin\Debug", "");
            path = path + @"App.config";
            XmlDocument doc = new XmlDocument();

            doc.Load(path);
            DataSet data = new DataSet();
            data.ReadXml(new XmlNodeReader(doc));

            conStr = data.Tables[3].Rows[0][1].ToString();
            return conStr;
        }
        public static void EjecutaQuerySql(string sql)
        {


            System.Data.SqlClient.SqlConnection cnn = new System.Data.SqlClient.SqlConnection(CadenaConexion());
            System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(sql, cnn);
            cmd.CommandType = CommandType.Text;

            cnn.Open();
            cmd.ExecuteNonQuery();
            cnn.Close();

        }
    }
}
