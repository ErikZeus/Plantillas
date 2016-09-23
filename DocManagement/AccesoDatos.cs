using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Collections;
using System.Text;
using System.Data.OleDb;
using System.Diagnostics;
using Microsoft.VisualBasic;
using System.Security.Cryptography;
using System.Configuration;
using System.IO;
using System.Web.Configuration;
using System.Net.Mail;
using System.Xml;
using Npgsql;
using MySql;

namespace Cotizador
{
    public class AccesoDatos
    {
        public string  EjecutaQueryNpgsql(string _sql)
        {

            string msg = "";
            try
            {
                NpgsqlConnection cnn = new NpgsqlConnection(ConfigurationManager.ConnectionStrings["cnnPostgres"].ToString());
                // Define a query
                NpgsqlCommand cmd = new NpgsqlCommand(_sql, cnn);
                cmd.CommandType = CommandType.Text;
                cnn.Open();
                cmd.ExecuteNonQuery();
                cnn.Close();
                return "Terminado";
            }
            catch (Exception ex)
            {
                return ex.Message.ToString() ;
            }
        }
        public string EjecutaQueryRegresaCadenaNpgsql(string _sql)
        {
            DataTable content = new DataTable();
            string retorna = "";

            try
            {
                NpgsqlConnection cnn = new NpgsqlConnection(ConfigurationManager.ConnectionStrings["cnnPostgres"].ToString());
                // Define a query
                NpgsqlCommand cmd = new NpgsqlCommand(_sql, cnn);
                NpgsqlDataAdapter adpt = new NpgsqlDataAdapter(cmd);

                cmd.CommandType = CommandType.Text;
                cnn.Open();
                adpt.Fill(content);
                cnn.Close();

                foreach (DataRow rw in content.Rows)
	            {
                    retorna = rw[0].ToString();
                    break;
	           }


                return retorna;
            }
            catch (Exception ex)
            {
              return  "Error: " + ex.Message.ToString();
            }
        }
        public DataTable  RegresaTablaNpqsql(string _sql)
        {
            DataTable content = new DataTable();
            string msg = "";
            try
            {
                NpgsqlConnection cnn = new NpgsqlConnection(ConfigurationManager.ConnectionStrings["cnnPostgres"].ToString());
                // Define a query
                NpgsqlCommand cmd = new NpgsqlCommand(_sql, cnn);
                NpgsqlDataAdapter adpt = new NpgsqlDataAdapter(cmd);

                cmd.CommandType = CommandType.Text;
                cnn.Open();
                adpt.Fill(content);
                cnn.Close();

                return content;
            }
            catch (Exception ex)
            {
                return content;
            }
    
          
        }
        public static string RegresaCadena_1_ResultadoMysql(string sql)
        {

            MySql.Data.MySqlClient.MySqlConnection cnn = new  MySql.Data.MySqlClient.MySqlConnection(ConfigurationManager.ConnectionStrings["cnnMysql"].ToString());

            MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand(sql, cnn);
            cmd.CommandType = CommandType.Text;
            MySql.Data.MySqlClient.MySqlDataAdapter adpt = new MySql.Data.MySqlClient.MySqlDataAdapter(cmd);
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
        public static DataTable RegresaTablaMySql(string sql)
        {

            MySql.Data.MySqlClient.MySqlConnection cnn = new MySql.Data.MySqlClient.MySqlConnection(ConfigurationManager.ConnectionStrings["cnnMysql"].ToString());

            MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand(sql, cnn);
            cmd.CommandType = CommandType.Text;
            MySql.Data.MySqlClient.MySqlDataAdapter adpt = new MySql.Data.MySqlClient.MySqlDataAdapter(cmd);
            System.Data.DataTable content = new System.Data.DataTable();

            cnn.Open();
            adpt.Fill(content);
            cnn.Close();

            return content;
        }

        public static void EjecutaQueryMySql(string sql)
        {

            MySql.Data.MySqlClient.MySqlConnection cnn = new MySql.Data.MySqlClient.MySqlConnection(ConfigurationManager.ConnectionStrings["cnnMysql"].ToString());
            MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand(sql, cnn);
            cmd.CommandType = CommandType.Text;

            cnn.Open();
            cmd.ExecuteNonQuery();
            cnn.Close();

        }


    }
}
public class Encript
{
    /// <summary>
    /// Encrypt a string using dual encryption method. Return a encrypted cipher Text
    /// </summary>
    /// <param name="toEncrypt">string to be encrypted</param>
    /// <param name="useHashing">use hashing? send to for extra secirity</param>
    /// <returns></returns>
    public static string Encriptar(string toEncrypt)
    {
        if (string.IsNullOrEmpty(toEncrypt))
            throw new ArgumentNullException("plainText");
        string outStr = null;
        // Encrypted string to return         
        RijndaelManaged aesAlg = null;
        // RijndaelManaged object used to encrypt the data.          
        try
        {
            // generate the key from the shared secret and the salt             
            Rfc2898DeriveBytes key = new Rfc2898DeriveBytes("MHTHJ", UTF8Encoding.UTF8.GetBytes("bpDYJrbpcM"));
            // Create a RijndaelManaged object             // with the specified key and IV.             
            aesAlg = new RijndaelManaged();
            aesAlg.Key = key.GetBytes(aesAlg.KeySize / 8);
            aesAlg.IV = key.GetBytes(aesAlg.BlockSize / 8);
            // Create a decrytor to perform the stream transform.             
            ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);
            // Create the streams used for encryption.             
            using (MemoryStream msEncrypt = new MemoryStream())
            {
                using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                {
                    using (StreamWriter swEncrypt = new StreamWriter(csEncrypt))
                    {
                        //Write all data to the stream.                         
                        swEncrypt.Write(toEncrypt);
                    }
                }
                outStr = Convert.ToBase64String(msEncrypt.ToArray());
            }
        }
        finally
        {
            // Clear the RijndaelManaged object.             
            if (aesAlg != null)
                aesAlg.Clear();
        }          // Return the encrypted bytes from the memory stream.         
        return outStr;
    }
    /// <summary>
    /// Get a parameter value
    /// </summary>
    /// <param name="name">Parameter name</param>
    /// <returns></returns>
    public static string TraerElvalorDelParametro(string name)
    {

        //return Encript.Decrypt(ConfigurationManager.AppSettings[name].ToString(), true);
        return Desencriptar(ConfigurationManager.AppSettings[name].ToString());

    }
    public static void GuardarElValorDelParametro(string llave, string valor)
    {
        string key = Encript.Encriptar(llave);
        string value = Encript.Encriptar(valor);
        Configuration webConfig = WebConfigurationManager.OpenWebConfiguration("~");
        webConfig.AppSettings.Settings.Remove(llave);
        webConfig.AppSettings.Settings.Add(llave, value);

        webConfig.Save();

    }
    /// <summary>
    /// DeCrypt a string using dual encryption method. Return a DeCrypted clear string
    /// </summary>
    /// <param name="cipherString">encrypted string</param>
    /// <param name="useHashing">Did you use hashing to encrypt this data? pass true is yes</param>
    /// <returns></returns>
    public static string Desencriptar(string llave)
    {
        bool useHashing = true;

        if (string.IsNullOrEmpty(llave))
            throw new ArgumentNullException("cipherText");

        // Declare the RijndaelManaged object         
        // used to decrypt the data.         
        RijndaelManaged aesAlg = null;
        // Declare the string used to hold        
        // the decrypted text.         
        string plaintext = null;
        try
        {
            // generate the key from the shared secret and the salt            
            Rfc2898DeriveBytes key = new Rfc2898DeriveBytes("MHTHJ", UTF8Encoding.UTF8.GetBytes("bpDYJrbpcM"));

            // Create a RijndaelManaged object             
            // with the specified key and IV.             
            aesAlg = new RijndaelManaged();
            aesAlg.Key = key.GetBytes(aesAlg.KeySize / 8);
            aesAlg.IV = key.GetBytes(aesAlg.BlockSize / 8);
            // Create a decrytor to perform the stream transform.             
            ICryptoTransform decryptor = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV);
            // Create the streams used for decryption.                             
            byte[] bytes = Convert.FromBase64String(llave);
            using (MemoryStream msDecrypt = new MemoryStream(bytes))
            {
                using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                {
                    using (StreamReader srDecrypt = new StreamReader(csDecrypt))
                        // Read the decrypted bytes from the decrypting stream                        
                        // and place them in a string.                         
                        plaintext = srDecrypt.ReadToEnd();
                }
            }
        }
        finally
        {             // Clear the RijndaelManaged object.             
            if (aesAlg != null)
                aesAlg.Clear();
        }
        return plaintext;
    }
}

public class Helper
{ 

    public static void RegistrarEvento(string EventoInfo)
        {
            string YearFolder = System.DateTime.Now.Year.ToString();
            string LogFolderDir = Path.Combine( System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + "/Logs", YearFolder);
            string LogFileName = Path.Combine(LogFolderDir, System.DateTime.Now.ToString("yyyyMMdd") + ".txt");
            FileStream fs = default(FileStream);
            StreamWriter sw = default(StreamWriter);

            try
            {
                if (!Directory.Exists(LogFolderDir))
                {
                    Directory.CreateDirectory(LogFolderDir);
                }

                if (!File.Exists(LogFileName))
                {
                    fs = File.Create(LogFileName);
                }
                else
                {
                    fs = new FileStream(LogFileName, FileMode.Append, FileAccess.Write);
                }

                sw = new StreamWriter(fs);

                sw.WriteLine("[{0}] - {1}", System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), EventoInfo);

                sw.Close();
                fs.Close();

            }
            catch (Exception ex)
            {
            }
        }

}