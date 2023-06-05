using Microsoft.Office.Interop.Excel;
using System.Data.OleDb;

namespace Muhasebe
{
    
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        public static string path = Directory.GetCurrentDirectory();
        [STAThread]
        static void Main()
        {
           
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
           //  System.Windows.Forms.Application.Run(new Form1());
            System.Windows.Forms.Application.Run(new Form1());

            ///////
            ///


        }
    }
}