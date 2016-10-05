/*
 * Author: Thilina Chandima
 * License : Free to Use Anybody
 * Date of Create : 26/03/2012
 * Time of Create : 09.09 PM
 * File Name : DBHandler.cs
 * Contact : tcgunarathena@gmail.com
 * 
 * NOTE: Appreciate Any Suggestions or Comments 
 * */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


 
    public class DBHandler
    {        
        public static string SrvName = @"192.168.5.5"; //Your SQL Server Name
        public static string DbName = @"seguro";//Your Database Name
        public static string UsrName = "sa";//Your SQL Server User Name
        public static string Pasword = "PromoAdmin12";//Your SQL Server Password
 
        /// <summary>
        /// Public static method to access connection string throw out the project 
        /// </summary>
        /// <returns>return database connection string</returns>
        public static string GetConnectionString()
        {
            return "Data Source=" + SrvName + "; initial catalog=" + DbName + "; user id=" + UsrName + "; password=" + Pasword + ";";//Build Connection String and Return
        }
    }
 







