/*
 
https://stackoverflow.com/questions/10591600/insert-datatable-into-excel-using-microsoft-access-database-engine-via-oledb
 Tek satýrda tanýmlarsan çalýþýyor, ayrý tanýmlarsan hata veriyor ????? > OleDbCommand commd = new OleDbCommand("SELECT * INTO Table1 FROM " + dt, conn);


 */namespace Muhasebe
{



    using Spire.Xls;
    using System;
    using System.ComponentModel;
    using System.Data;
    using System.Data.OleDb;
    using System.Diagnostics;
    using System.Reflection;
    using System.Text;
    using DataTable = System.Data.DataTable;

    using accInterop = Microsoft.Office.Interop.Access;
    using excelInterop = Microsoft.Office.Interop.Excel;
    //using Excel = Microsoft.Office.Interop.Excel;
    //using ExcelAdaptorLib;
    //using System.Runtime.InteropServices;
    //using Microsoft.Office.Interop.Excel;
    using Workbook = Spire.Xls.Workbook;
    using Microsoft.Office.Interop.Excel;
    using Microsoft.Office.Interop.Access.Dao;
    using static System.ComponentModel.Design.ObjectSelectorEditor;
    using System.Diagnostics.Metrics;
    using System.Windows.Forms;
    using Microsoft.VisualBasic.ApplicationServices;


    public partial class Form1 : Form
    {
        int int1;
        //  static string path = Directory.GetCurrentDirectory();
        static string path = "C:\\Muhasebe";
        Int32 rowCount;
        Muhasebe.Fonsiyonlar fonks = new Muhasebe.Fonsiyonlar();


        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            // Test2_Click(null, null);
            // btExportExcel_Click(null, null);  
            System.Globalization.CultureInfo customCulture = (System.Globalization.CultureInfo)System.Threading.Thread.CurrentThread.CurrentCulture.Clone();
            customCulture.NumberFormat.NumberDecimalSeparator = ".";
            System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;

            //btLink2Table.Enabled = true;
            //btMuhasebe2Banka.Enabled = false;
            //btBanka2Talimat.Enabled = false;
            //btExportExcel.Enabled = false;
            //btBanka2Talimat_Click(null, null);


        }
        private void btMuhasebe2Banka_Click(object sender, EventArgs e)
        {
            /*
             Muhasebe.Tutarý al, Banka.Tutarda ara. Tek kayýt bulursan Muhasebedeki fiþ nosunu al, boþ ise Banka.OracleMahsupNo'ya ve boþ ise Muhasebe.KontrolMahsup'a yaz.
            */



            progressBar1.Value = 0;
            progressBar1.Step = 1;
            progressBar1.Minimum = 0;
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.RunWorkerAsync();

            //string path = Directory.GetCurrentDirectory();
            System.Globalization.CultureInfo customCulture = (System.Globalization.CultureInfo)System.Threading.Thread.CurrentThread.CurrentCulture.Clone();
            customCulture.NumberFormat.NumberDecimalSeparator = ".";
            System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;

            OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + "\\Muhasebe.accdb");
            OleDbCommand cmd = con.CreateCommand();
            OleDbCommand cmdBanka = con.CreateCommand();
            //cmd.CommandText = "select * from Muhasebe where string.IsNullOrEmpty('Kontrol (Mahsup)')";
            //cmd.CommandText = "select * from Muhasebe where (isnull(@'Kontrol (Mahsup)', '') = '' )";
            //cmd.CommandText = "select  Muhasebe.[Kontrol (Mahsup)] from Muhasebe where 'Kontrol (Mahsup)' = ''";
            //cmd.CommandText = "select  Muhasebe.[Kontrol (Mahsup)] from Muhasebe where 'Kontrol (Mahsup)' = dbnull";
            //cmd.CommandText = "select ID, Muhasebe.[KontrolMahsup], [FisNo], Tutar from Muhasebe where [KontrolMahsup]  IS NULL or trim([KontrolMahsup])=''";
            cmd.CommandText = "select  ID, [FisNo], Tutar from Muhasebe where [KontrolMahsup]  IS NULL or trim([KontrolMahsup])=''";
            //SELECT COUNT(DISTINCT Country) FROM Customers;
            // SELECT Count(*) AS DistinctCountries
            // FROM(SELECT DISTINCT Country FROM Customers);
            cmd.Connection = con;
            con.Open();
            DataTable dtMuhasebe = new DataTable();
            DataTable dtBanka = new DataTable();
            OleDbDataAdapter adMuhasebe = new OleDbDataAdapter();

            adMuhasebe.SelectCommand = cmd;
            adMuhasebe.Fill(dtMuhasebe);
            rowCount = dtMuhasebe.Rows.Count;
            progressBar1.Maximum = rowCount;
            //   MessageBox.Show(rowCount.ToString());
            lbBilgi.Text = "Bulunan satýr sayýsý: " + rowCount.ToString() + ".\t Muhasebe > Banka iþlemi devam ediyor...";
            Cursor.Current = Cursors.WaitCursor;
            // dtMuhasebe                = dtMuhasebe.DefaultView.ToTable(true, "Tutar"); <- Bu tek kolonu alýyor.
            foreach (DataRow rowMuhasebe in dtMuhasebe.Rows)
            {
              
                OleDbCommand cmd3 = con.CreateCommand();
                cmd3.CommandText = "select count(Tutar) from Muhasebe where Tutar=" + rowMuhasebe["Tutar"];
                // cmd3.CommandText = "select count(Tutar) from "+ dtMuhasebe +" where Tutar=" + rowMuhasebe["Tutar"];
                int? tekrar = (int)cmd3.ExecuteScalar();
                if (tekrar > 1)
                {
                    progressBar1.PerformStep();
                    continue;
                }



                double? dblAranan = 0.00;
                if (rowMuhasebe["Tutar"] != DBNull.Value)
                {
                    dblAranan = (double)rowMuhasebe["Tutar"];
                }
                else
                {
                    progressBar1.PerformStep();
                    continue;
                }
                cmdBanka.CommandText = "select * from Banka where Banka.[Tutar]=" + dblAranan;
                OleDbDataAdapter daBanka = new OleDbDataAdapter();
                dtBanka.Clear();
                daBanka.SelectCommand = cmdBanka;
                //  cmdBanka.ExecuteNonQuery();

                daBanka.Fill(dtBanka);




                //   adMuhasebe.SelectCommand = cmd;

                int1 = dtBanka.Rows.Count;
            //    MessageBox.Show("kayýt sayýsý : " + int1.ToString() + "tutar : " + rowMuhasebe["Tutar"].ToString() + "\n" + dtMuhasebe.Rows[0]["FisNo"].ToString());
                if (int1 == 1)
                {
                    DataRow drBanka = dtBanka.Rows[0];
                    int idBanka = (int)drBanka["ID"];
                    int idMuhasebe = (int)rowMuhasebe["ID"];
                    //   drBanka["OracleMahsupNo"] = 1869;
                    if (dtMuhasebe.Rows[0]["FisNo"] != DBNull.Value)
                    {
                    
                        double? FisNo = (double)dtMuhasebe.Rows[0]["FisNo"];
                        //  cmdBanka = new OleDbCommand("update Banka set OracleMahsupNo=IsNull(Fisno, " + FisNo + ") where ID=" + idBanka);
                        cmdBanka = new OleDbCommand("update Banka set OracleMahsupNo=" + rowMuhasebe["FisNo"] + ", MuhasebeTalimat='Muhasebe' where ID=" + idBanka + " and (OracleMahsupNo) Is Null");

                        daBanka.UpdateCommand = cmdBanka;
                        daBanka.UpdateCommand.Connection = con;
                        // daBanka.UpdateCommand.Connection.Open();
                        daBanka.UpdateCommand.ExecuteNonQuery();

                        // cmd = new OleDbCommand("update Muhasebe set KontrolMahsup=" +  FisNo + " where ID=" + idMuhasebe + " and (KontrolMahsup) Is Null");
                        cmd = new OleDbCommand("update Muhasebe set KontrolMahsup=" + rowMuhasebe["FisNo"] + ", MuhasebeTalimat='Muhasebe' where ID=" + idMuhasebe + " and (KontrolMahsup) Is Null");
                        adMuhasebe.UpdateCommand = cmd;
                        adMuhasebe.UpdateCommand.Connection = con;
                        // daBanka.UpdateCommand.Connection.Open();
                        adMuhasebe.UpdateCommand.ExecuteNonQuery();
                        dtBanka.Clear();
                        daBanka.Dispose();
                    }



                    // MessageBox.Show(drBanka.RowState.ToString());
                    //dr.SetModified();



                    //  dtBanka.Rows[0]["OracleMahsupNo"] = row["FisNo"];




                    //     
                    //    dtBanka.AcceptChanges();
                    // daBanka.Update(dtBanka);
                    //   MessageBox.Show(dtBanka.Rows[0]["OracleMahsupNo"].ToString());



                    dtBanka.Clear();
                    daBanka.Dispose();
                }


                //foreach (DataColumn column in dtBanka.Columns)
                //{
                //  //string  urunAdi = row["UrunAdi"].ToString();
                //  //  urunAciklamasi = row["UrunAciklamasi"].ToString();
                //}
                progressBar1.PerformStep();

            }

            adMuhasebe.Dispose();
            con.Close();
            btLink2Table.Enabled = false;
            btMuhasebe2Banka.Enabled = false;
            btBanka2Talimat.Enabled = true;
            btExportExcel.Enabled = false;
            this.Cursor = Cursors.Arrow;
            lbBilgi.Text = ("Ýþlem Tamamlandý...");
           // MessageBox.Show("iþlem Tamamlandý.");
            lbBilgi.Text = "Muhasebe > Banka iþlemi bitti...";
           // btBanka2Talimat_Click(this, null);   

        }



        private void btExportExcel_Click(object sender, EventArgs e)
        {
            //  Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            // workbook.LoadFromFile(@path + "\\MuhasebeExport.xlsx");

            Database2Excel();
            btLink2Table.Enabled = true;
            btMuhasebe2Banka.Enabled = false;
            btBanka2Talimat.Enabled = false;
            btExportExcel.Enabled = false;
        }

        private void btBanka2Talimat_Click(object sender, EventArgs e)
        {
            /*
             banka.tutarda ki - olan deðerleri talimat.tutar'da ara.
            1 adet kayýt bulursan talimat.BankaTalimat
'yi true yap
            Talimat.HesapNo'yu Banka.MuhasebeKodu'na yaz.
             */
            //  string path = Directory.GetCurrentDirectory();
            progressBar1.Value = 0;
            progressBar1.Step = 1;
            progressBar1.Minimum = 0;
           
          //  backgroundWorker1.Dispose();
          //  backgroundWorker1.WorkerReportsProgress = true;
          //  backgroundWorker1.RunWorkerAsync();

            DataSet dSet;
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + "\\Muhasebe.accdb");
            OleDbCommand cmdBanka = con.CreateCommand();
            OleDbCommand cmdTalimat = con.CreateCommand();

            cmdBanka.CommandText = "select ID, Tutar, OracleMahsupNo from Banka where Tutar < 0 and MuhasebeKodu IS NULL or trim([MuhasebeKodu])=''";
            cmdBanka.Connection = con;
            con.Open();
            DataTable dtBanka = new DataTable();
            DataTable dtTalimat = new DataTable();
            OleDbDataAdapter daBanka = new OleDbDataAdapter();
            OleDbDataAdapter adTalimat = new OleDbDataAdapter();
            daBanka.SelectCommand = cmdBanka;
            daBanka.Fill(dtBanka);
            rowCount = dtBanka.Rows.Count;
            progressBar1.Maximum = rowCount;
            lbBilgi.Text = "Bulunan satýr sayýsý: " + rowCount.ToString() + ".\t Muhasebe > Talimat iþlemi devam ediyor...";
            // MessageBox.Show("Bulunan satýr sayýsý: " + rowCount.ToString());

            DataTable dt = dtBanka.DefaultView.ToTable(true, "Tutar");
            Cursor.Current = Cursors.WaitCursor;
            foreach (DataRow rowBanka in dtBanka.Rows)
            {
                double? dblAranan = (double)rowBanka["Tutar"];
                //////

                OleDbCommand cmd3 = con.CreateCommand();
                cmd3.CommandText = "select count(Tutar) from Muhasebe where Tutar=" + rowBanka["Tutar"];
                // cmd3.CommandText = "select count(Tutar) from "+ dtMuhasebe +" where Tutar=" + rowMuhasebe["Tutar"];
                int? tekrar = (int)cmd3.ExecuteScalar();
                if (tekrar > 1)
                {
                    progressBar1.PerformStep();
                    continue;
                }
                /////////
                cmdTalimat.CommandText = "select ID, Tutar, HesapNo from Talimat where Talimat.[Tutar]=" + dblAranan;
                adTalimat.SelectCommand = cmdTalimat;
                dtTalimat.Clear();
                adTalimat.Fill(dtTalimat);

                //   daBanka.SelectCommand = cmd;

                int1 = dtTalimat.Rows.Count;

              //  MessageBox.Show("kayýt sayýsý : " + int1.ToString() + "tutar : " + rowBanka["Tutar"].ToString() + "\n" + rowBanka["FisNo"].ToString());
                if (int1 == 1)
                {
                    DataRow drTalimat = dtTalimat.Rows[0];
                    int idTalimat = (int)drTalimat["ID"];
                    int idBanka = (int)rowBanka["ID"];
                    //   drTalimat["OracleMahsupNo"] = 1869;
                    if (dtBanka.Rows[0]["OracleMahsupNo"] == DBNull.Value)
                    {
                        // string? FisNo = (string)dtBanka.Rows[0]["FisNo"];
                        cmdTalimat = new OleDbCommand("update Talimat set BankaTalimat='Talimat' where ID=" + idTalimat);
                        adTalimat.UpdateCommand = cmdTalimat;
                        adTalimat.UpdateCommand.Connection = con;
                        // adTalimat.UpdateCommand.Connection.Open();
                        adTalimat.UpdateCommand.ExecuteNonQuery();
                       
                        cmdBanka = new OleDbCommand("update Banka set MuhasebeKodu='" + drTalimat["HesapNo"].ToString() + "', BankaTalimat='Talimat' where ID=" + idBanka);
                        daBanka.UpdateCommand = cmdBanka;
                        daBanka.UpdateCommand.Connection = con;
                        // adTalimat.UpdateCommand.Connection.Open();
                        daBanka.UpdateCommand.ExecuteNonQuery();
                    }



                    // MessageBox.Show(drTalimat.RowState.ToString());
                    //dr.SetModified();



                    //  dtTalimat.Rows[0]["OracleMahsupNo"] = row["FisNo"];




                    //     
                    //    dtTalimat.AcceptChanges();
                    // adTalimat.Update(dtTalimat);
                    //   MessageBox.Show(dtTalimat.Rows[0]["OracleMahsupNo"].ToString());





                }

                progressBar1.PerformStep();

                //foreach (DataColumn column in dtTalimat.Columns)
                //{
                //  //string  urunAdi = row["UrunAdi"].ToString();
                //  //  urunAciklamasi = row["UrunAciklamasi"].ToString();
                //}

            }
            adTalimat.Dispose();
            daBanka.Dispose();
            con.Close();
            btLink2Table.Enabled = false;
            btMuhasebe2Banka.Enabled = false;
            btBanka2Talimat.Enabled = false;
            btExportExcel.Enabled = true;
            this.Cursor = Cursors.Arrow;
            lbBilgi.Text = ("Ýþlem Tamamlandý...");
            //MessageBox.Show("iþlem Tamamlandý.");
            lbBilgi.Text = ".\t Muhasebe > Talimat iþlemi bitti.";
          //  btExportExcel_Click(this, new EventArgs()); 

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










        public void Excel2Database()
        {/*
          dt'den veritabanýna yazdýramadým...
          */
            Cursor.Current = Cursors.WaitCursor;
            OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + "\\Muhasebe.accdb");
            OleDbDataAdapter da = new OleDbDataAdapter();
            OleDbCommand cmd = new OleDbCommand();
            progressBar1.Value = 0;
            progressBar1.Step = 1;
            progressBar1.Minimum = 0;
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.RunWorkerAsync();
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(path + @"\\Muhasebe.xlsx");
            DataSet ds = new DataSet();
            conn.Open();
            DataTable dt = new DataTable();
            for (int i = 0; i < 1; i++)
            {
                // ID kolonlarýný autonumber yap unique index, primary key yapma
                Spire.Xls.Worksheet sheet = workbook.Worksheets[i];
                string tabloAdi = workbook.Worksheets[i].Name;
                lbBilgi.Text = ("Excel'den " + tabloAdi + " tablosuna yazýlýyor...");
                // da.Fill(dt);
                dt = sheet.ExportDataTable();
                cmd.CommandText = "Insert INTO " + tabloAdi + " SELECT TOP 10 * FROM " + dt;
                // cmd.CommandText = "select * INTO " + tabloAdi + "xxx FROM Talimat";çalýþtý

                cmd.CommandText = "select * INTO " + tabloAdi + " FROM " + dt;
                cmd.Connection = conn;
                OleDbCommand commd = new OleDbCommand("SELECT * INTO Table1 FROM " + dt, conn);
                // da.InsertCommand= cmd;
                //  da.UpdateBatchSize
                //  da.Update(dt);

                //  dt.Columns.RemoveAt(0);

                //cmd.CommandText = "DROP INDEX Kimlik ON Talimat"; 
                //cmd.ExecuteNonQuery();
                //cmd.CommandText= "ALTER TABLE Talimat DROP COLUMN ID" ;
                //cmd.ExecuteNonQuery();
                dt.Clone();
                commd.ExecuteNonQuery();



                //cmd.CommandText = "ALTER TABLE Talimat ADD ID INT NOT NULL AUTO_INCREMENT";
                //cmd.ExecuteNonQuery();
                //cmd.CommandText = "CREATE UNIQUE INDEX Kimlik ON Talimat(ID)";
                //cmd.ExecuteNonQuery();
                progressBar1.PerformStep();
                lbBilgi.Text = ("Excel'den " + tabloAdi + " tablosuna yazýldý...");


            }
            cmd.Dispose();
            dt.Dispose();
            conn.Close();
            conn.Dispose();


            Cursor.Current = Cursors.Arrow;
            progressBar1.Maximum = rowCount;
            progressBar1.PerformStep();
            lbBilgi.Text = ("Excel'den veritabanýna yazma iþlemi bitti.");
            MessageBox.Show("Bitti.");
            lbBilgi.Text = "...";
        }

        private void VeritabaninaYaz_Click(object sender, EventArgs e)
        {
            Excel2Database();
        }

        private void Link2Table_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            string[] tablolar = { "Talimat", "Banka", "Muhasebe" };
            //  string tabloAdi = tablolar[0];
            OleDbConnection conn = fonks.OpenConn();
            OleDbCommand cmd = new OleDbCommand();

            var schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            cmd.Connection = conn;
            foreach (string str1 in tablolar)
            {
                foreach (var row in schema.Rows.OfType<DataRow>())
                {
                    if (row.ItemArray[2].ToString() == str1)
                    {
                        cmd.CommandText = "DROP TABLE " + str1;
                        cmd.ExecuteNonQuery();
                        break;
                    }
                }
            }
            foreach (string tabloAdi in tablolar)
            {
                OleDbCommand command = new OleDbCommand("select * into " + tabloAdi + " from Link" + tabloAdi, conn);
                cmd.CommandText = "select * into " + tabloAdi + " from Link" + tabloAdi;
                cmd.ExecuteNonQuery();

                //   cmd.CommandText = "CREATE UNIQUE INDEX Kimlik ON "+ tabloAdi+"(ID)";
                cmd.CommandText = "ALTER TABLE " + tabloAdi + " ADD COLUMN ID AutoIncrement";
                cmd.ExecuteNonQuery();
            }
            conn.Close();
            btLink2Table.Enabled = false;
            btMuhasebe2Banka.Enabled = true;
            btBanka2Talimat.Enabled = false;
            btExportExcel.Enabled = false;
            Cursor.Current = Cursors.Arrow;
            lbBilgi.Text = "Excel tablolarý import edildi.";
            // MessageBox.Show("bitti.");
          //  btMuhasebe2Banka_Click(null, null);


        }

        private void btKapat_Click(object sender, EventArgs e)
        {
            // ExcelAc();
            Close();
        }

        public void Database2Excel()
        {
            Cursor.Current = Cursors.WaitCursor;
            OleDbConnection conn = fonks.OpenConn();
            OleDbCommand cmd = new OleDbCommand("alter table Talimat drop column ID", conn);
            cmd.ExecuteNonQuery();
            cmd.CommandText = "alter table Banka drop column ID";
            cmd.ExecuteNonQuery();
            cmd.CommandText = "alter table Muhasebe drop column ID";
            cmd.ExecuteNonQuery();
            conn.Close();
            conn.Dispose();
            cmd.Dispose();


            DataTable dt = new DataTable();
            Spire.Xls.Workbook workbook = new Workbook();
            workbook.Worksheets.Clear();
            dt = GetDataTableFromDB("select * from Muhasebe");
            Spire.Xls.Worksheet worksheet = workbook.Worksheets.Add("Muhasebe");
            worksheet.InsertDataTable(dt, true, 1, 1, true);

            dt = GetDataTableFromDB("select * from Banka");
            worksheet = workbook.Worksheets.Add("Banka");
            worksheet.InsertDataTable(dt, true, 1, 1, true);

            dt = GetDataTableFromDB("select * from Talimat");
            worksheet = workbook.Worksheets.Add("Talimat");
            worksheet.InsertDataTable(dt, true, 1, 1, true);
            string tarih = DateTime.Now.ToString();
            tarih = tarih.Replace(":", "_");
            tarih = tarih.Replace(".", "_");
            tarih = tarih.Replace(" ", "_");
            workbook.SaveToFile(path + "\\Export\\MuhasebeExport" + tarih + ".xlsx", ExcelVersion.Version2016);

            Cursor.Current = Cursors.Arrow;
            MessageBox.Show("Ýþlem bitti.");

        }
        public void ExcelAc()



        {
            //Process p = new Process();
            //p.StartInfo = new ProcessStartInfo()
            //{
            //    CreateNoWindow = true,               
            //    FileName = "Excel.exe"
            //};
            //p.Start();


            //// excelInterop.Workbook  worKbooK = excelInterop.Workbooks.Add(Type.Missing);
            //Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            ////  Process.Start("Excel.exe", @path + "\\MuhasebeExport.xlsx");
            //System.Diagnostics.Process.Start("E:\\Kodlama\\_MelihProgramlar\\_WindowsForms\\Muhasebe\\Muhasebe\\bin\\Debug\\net7.0-windows\\MuhasebeExport.xlsx");


            // excelInterop.Application xxx = new excelInterop.Application();
            //excelInterop.Application.Workbooks.Open(@"C:\Test\YourWorkbook.xlsx");
            // excelInterop.Workbook wb = xxx.Workbooks.Open(path+ "\\MuhasebeExport.xlsx");
            //excelInterop.Application xlApp;
            // excelInterop.Workbook xlWorkBook;
            // excelInterop.Worksheet xlWorkSheet;
            // object misValue = System.Reflection.Missing.Value;

            // xlApp = new excelInterop.ApplicationClass;
            // xlWorkBook = xlApp.Workbooks.Open("csharp.net-informations.xls", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            // xlWorkSheet = (excelInterop.Worksheet)xlWorkBook.Worksheets.get_Item(1);
        }
        public void xxxx()
        {


        }


        #region ######################################################
        /*
                public void DataTableToWorkBook(DataTable dt)
                {
                    object oMissing = Missing.Value;

                    Microsoft.Office.Interop.Excel.Application app;
                    Microsoft.Office.Interop.Excel.Workbooks wkBks;
                    Microsoft.Office.Interop.Excel.Workbook wkBk;
                    Microsoft.Office.Interop.Excel.Sheets wkShts;
                    Microsoft.Office.Interop.Excel.Worksheet wkSht;



                    string fileLink;
                    string filename = path + "\\MuhasebeSonRapor.xls";
                    string sheetName = "mmmm";
                    FileStream fs = new FileStream(filename, FileMode.Create);
                    fs.Close();

                    app = new Microsoft.Office.Interop.Excel.Application();
                    try
                    {
                        app.DisplayAlerts = true;
                        wkBks = app.Workbooks;
                        wkBk = wkBks.Open(filename, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                        wkShts = (Microsoft.Office.Interop.Excel.Sheets)wkBk.Sheets;
                        wkSht = (Microsoft.Office.Interop.Excel.Worksheet)wkShts.get_Item(1);






           //              * 
           //              *  Excel.Worksheet newWorksheet;
           //             newWorksheet = (Excel.Worksheet)wkBk.Worksheets.Add();
           //             newWorksheet.Name = "Muhasebe";
           //              * 
           //wkSht = wkBk.Sheets[1];

           //             Microsoft.Office.Interop.Excel.Worksheet aaa  = (Worksheet)app.Worksheets["Sheet1"];
           //             aaa.Name = "NewTabName";



                        //  MessageBox.Show(wkSht.Name+ (Microsoft.Office.Interop.Excel.Worksheet)wkShts.get_Item(1).ToString());
                        #region Captions
                        for (var i = 0; i < dt.Columns.Count; i++)
                        {
                            //   Range row1 =  wkSht.Rows.Cells[1, 1];
                            //  var cell = wkSht.Cells[1, i + 10];

                            // MessageBox.Show((string)cell);
                            //  cell.Value = dt.Columns[i].Caption;

                            wkSht.Cells[1, i + 1].Value = dt.Columns[i].Caption;
                        }
                        #endregion

                        #region Data

                        for (var r = 0; r < dt.Rows.Count; r++)
                            for (var c = 0; c < dt.Columns.Count; c++)
                            {
                                var cell = wkSht.Cells[r + 2, c + 1];
                                cell.Value = dt.Rows[r][c];
                            }

                        #endregion

                        app.Visible = true;


                        //  wkBk.Save();
                        

                            //            wkBk.Sheets.Add(After: wkBk.Sheets[wkBk.Sheets.Count]); wkBk.Worksheets.Add(
                            //System.Reflection.Missing.Value,
                            //wkBk.Worksheets[wkBk.Worksheets.Count],
                            //1,
                            //System.Reflection.Missing.Value);
                        

        wkBk.Close(true, oMissing, oMissing);

                Marshal.ReleaseComObject(wkSht);
                Marshal.ReleaseComObject(wkShts);
                Marshal.ReleaseComObject(wkBk);
                Marshal.ReleaseComObject(wkBks);

            }
            catch { }
            finally
            {
                app.Quit();

                Marshal.ReleaseComObject(app);

                wkSht = null;
                wkBk = null;
                app = null;
                GC.Collect();


            }
            MessageBox.Show("bitti");


        }
        private void Test1()
        {
            //   //connect database
            //   OleDbConnection connection = new OleDbConnection();
            ////   connection.ConnectionString @"Provider=""Microsoft.Jet.OLEDB.4.0"";Data Source=""demo.mdb"";User Id=;Password="
            //   OleDbCommand command = new OleDbCommand();
            //   command.CommandText = "select * from parts";
            //   DataSet dataSet = new System.Data.DataSet();
            //   OleDbDataAdapter dataAdapter = new OleDbDataAdapter(command.CommandText, connection);
            //   dataAdapter.Fill(dataSet);
            //   DataTable t = dataSet.Tables[0];
            //export datatable to excel
            //   Workbook book = new Workbook();
            //   Worksheet sheet = book.Worksheets[0];
            ////   sheet.data.InsertDataTable(t, true, 1, 1);
            //   book.SaveAs("insertTableToExcel.xls");
            //   System.Diagnostics.Process.Start("insertTableToExcel.xls");
        }
        private void Test2()
        {

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();



            if (xlApp == null)

            {

                MessageBox.Show("Excel is not properly installed!!");

                return;

            }





            Excel.Workbook xlWorkBook;

            Excel.Worksheet xlWorkSheet;

            object misValue = System.Reflection.Missing.Value;



            xlWorkBook = xlApp.Workbooks.Add(misValue);

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);



            xlWorkSheet.Cells[1, 1] = "ID";

            xlWorkSheet.Cells[1, 2] = "Name";

            xlWorkSheet.Cells[2, 1] = "1";

            xlWorkSheet.Cells[2, 2] = "One";

            xlWorkSheet.Cells[3, 1] = "2";

            xlWorkSheet.Cells[3, 2] = "Two";







            xlWorkBook.SaveAs(path + "\\csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

            xlWorkBook.Close(true, misValue, misValue);

            xlApp.Quit();



            Marshal.ReleaseComObject(xlWorkSheet);

            Marshal.ReleaseComObject(xlWorkBook);

            Marshal.ReleaseComObject(xlApp);



            MessageBox.Show("Excel file created , you can find the file d:\\csharp-Excel.xls");

        }
        private void Test2_Click(object sender, EventArgs e)
        {


            DataTableToWorkBook(GetDataTableFromDB("select top 10 * from Muhasebe"));
        }

        static void ExcelKaydet()
        {
            //  Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook workbook;
            Microsoft.Office.Interop.Excel.Worksheet worksheet;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            Excel.Application objExcel = new Excel.Application();
            // objExcel.Visible = true;
            workbook = objExcel.Workbooks.Add(Type.Missing);


            // worKbooK = Excel.Workbooks.Add(Type.Missing);
            // Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            //Process.Start("Excel.exe", @path + "\\MuhasebeExport.xls");
            //     System.Diagnostics.Process.Start(path + "\\MuhasebeExport.xls");


            //   Microsoft.Office.Interop.Excel.Workbook excelWorkBook = excelApp.Workbooks.Add();

            // Workbook wb = Microsoft.Office.Interop.Excel.Workbooks.Open(path + "\\MuhasebeExport.xlsx");

            //  Workbook wb = Microsoft.Office.Interop.Excel.Workbooks.Add(System.Reflection.Missing.Value);
            //   Microsoft.Office.Interop.Excel.Workbook workbook = new Micr
            //
            //   osoft.Office.Interop.Excel.Workbook();
            DataTable dt = GetDataTableFromDB("select top 1 * from Muhasebe");
            //Excel.Worksheet Page2;
            //Excel.Worksheet Page3;
            //   workbook.Worksheets.Add("WriteToCell");



            //     Worksheet sheet = workbook.Worksheets[0];
            //copy data from sheet into a datatable
            //    DataTable dataTable = sheet.ExportDataTable(dt,true,1,1);
            //load sheet1
            //  Worksheet sheet1 = workbook.Worksheets[workbook.ActiveSheet];
           
        worksheet = null;
            worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets["Sheet1"];
            worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
            worksheet.Name = "Exported from DataTable";
            Excel.Worksheet newWorksheet;
            newWorksheet = (Excel.Worksheet)workbook.Worksheets.Add();
            newWorksheet.Name = "lþlþþlþl";
            worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);
            workbook.Worksheets[1].InsertDataTable(dt, true, 1, 1);


            dt = GetDataTableFromDB("select top 1 * from Banka");
            workbook.Worksheets[1].InsertDataTable(dt, true, 1, 1);
            dt = GetDataTableFromDB("select top 1 * from Talimat");
            workbook.Worksheets[2].InsertDataTable(dt, true, 1, 1);
            //  workbook.SaveAs2(path + "\\MuhasebeExport.xlsx");
            //  System.Diagnostics.Process.Start(path + "\\MuhasebeExport.xlsx");

        }
        */
        #endregion ###################################################
    }
}