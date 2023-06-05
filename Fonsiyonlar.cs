using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Muhasebe
{
    internal class Fonsiyonlar
    {
        //public static string path = Directory.GetCurrentDirectory();
        static string path = "C:\\Muhasebe";
        public OleDbConnection OpenConn()
        {
            OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + "\\Muhasebe.accdb");            
              conn.Open();            
            return conn;
        }
    }
}
