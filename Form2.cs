using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;



namespace Muhasebe
{
    public partial class Form2 : Form
    {
        static string path = Directory.GetCurrentDirectory();
        public Form2()
        {
            InitializeComponent();
        }

        private void ExceleYaz_Click(object sender, EventArgs e)
        {
            Yaz1();
        }
        static void YazOrg()
        {
            //Create a Workbook instance
            Workbook workbook = new Workbook();
            //Remove default worksheets
            workbook.Worksheets.Clear();
            //Add a worksheet and name it
            Worksheet worksheet = workbook.Worksheets.Add("InsertDataTable");
            //Create a DataTable object
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("No", typeof(Int32));
            dataTable.Columns.Add("Name", typeof(String));
            dataTable.Columns.Add("City", typeof(String));
            //Create rows and add data
            DataRow dr = dataTable.NewRow();
            dr[0] = 1;
            dr[1] = "Tom";
            dr[2] = "New York";
            dataTable.Rows.Add(dr);
            dr = dataTable.NewRow();
            dr[0] = 2;
            dr[1] = "Jerry";
            dr[2] = "Huston";
            dataTable.Rows.Add(dr);
            dr = dataTable.NewRow();
            dr[0] = 3;
            dr[1] = "Dave";
            dr[2] = "Florida";
            dataTable.Rows.Add(dr);
            //Write datatable to the worksheet
            worksheet.InsertDataTable(dataTable, true, 1, 1, true);
            //Save to an Excel file
            workbook.SaveToFile(path + "\\InsertDataTable.xlsx", ExcelVersion.Version2016);
        }
        static void Yaz1()
        {
            Cursor.Current = Cursors.WaitCursor;
            DataTable dt = new DataTable();
            Workbook workbook = new Workbook();
            workbook.Worksheets.Clear();
            dt = GetDataTableFromDB("select top 1 * from Muhasebe");
            Worksheet worksheet = workbook.Worksheets.Add("Muhasebe");
            worksheet.InsertDataTable(dt, true, 1, 1, true);

            dt = GetDataTableFromDB("select top 1 * from Banka");
            worksheet = workbook.Worksheets.Add("Banka");
            worksheet.InsertDataTable(dt, true, 1, 1, true);

            dt = GetDataTableFromDB("select top 1 * from Talimat");
            worksheet = workbook.Worksheets.Add("Talimat");
            worksheet.InsertDataTable(dt, true, 1, 1, true);
            workbook.SaveToFile(path + "\\InsertDataTable.xlsx", ExcelVersion.Version2016);
            Cursor.Current = Cursors.Arrow;
            MessageBox.Show("bitti.");
        }

        static DataTable GetDataTableFromDB(string cmdText)
        {
            DataTable dt = new DataTable();


            using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + "\\Muhasebe.accdb"))
            {
                OleDbCommand commd = new OleDbCommand(cmdText, conn);
                using (OleDbDataAdapter da = new OleDbDataAdapter(commd))
                {
                    da.Fill(dt);
                }
            }
            return dt;
        }


        static void Excel2Database()
        {
            Cursor.Current = Cursors.WaitCursor;
            //Create a new workbook
            Workbook workbook = new Workbook();
            //Load a file and imports its data
            workbook.LoadFromFile(path + @"\\Muhasebe.xlsx");
            //Initialize worksheet
            Worksheet sheet = workbook.Worksheets[0];
            // get the data source that the grid is displaying data for
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            //  ds.Tables[0]..dataso
            //  this.dataGridView1.DataSource = sheet.ExportDataTable();
            dt = sheet.ExportDataTable();

            using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + "\\Muhasebe.accdb"))
            {
                OleDbCommand commd = new OleDbCommand("SELECT * INTO Table1 FROM " + dt, conn);
                conn.Open();
                commd.ExecuteNonQuery();

            }
            Cursor.Current = Cursors.Arrow;
            MessageBox.Show("bitti.");
        }

        private void VeritabaninaYaz_Click(object sender, EventArgs e)
        {
            Excel2Database();
        }
    }
}





