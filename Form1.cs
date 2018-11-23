using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Data.Common;
using System.Data.SqlClient;
using System.Data.OleDb;
using ClosedXML.Excel;
using MsgBox;
using System.IO;

namespace automateReport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
          
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            this.MaximizeBox = false;
            btExport.Enabled = false;
            txtServerName.Text = @"DESKTOP-MPK8RUP\SQLEXPRESS";
            txtDatabaseName.Text = "TestSecur";
            txtID.Text = "*********";
            txtPass.Text = "*********";
            txtOutputTable.Text = "CAL";
            txtCalTable.Text = "ExcelODS";
        }

        private static DataTable dt = new DataTable();//dataTable for add to dataSet
        public static bool finish = false; //----------DoSomething
        public static SaveFileDialog sfd = new SaveFileDialog();//dialog for select path to savefile
        public string fileName = null;//path string from "sfd"
        public string savePath = null;//only rootpath
        public string path = null;
        public static DataSet ds = new DataSet();// dataset for add to excel with closeXML
        public static Thread t1, t2;//loading thread

        #region **Connect Database 
        public void loading()
        {
            txtStatus.BackColor = Color.IndianRed;

            do
            {
                Thread.Sleep(500);
                backgroundWorker1.ReportProgress(0, "loading.");
                Thread.Sleep(500);
                backgroundWorker1.ReportProgress(0, "loading..");
                Thread.Sleep(500);
                backgroundWorker1.ReportProgress(0, "loading...");
                Thread.Sleep(500);
                backgroundWorker1.ReportProgress(0, "loading....");
                Thread.Sleep(500);
                backgroundWorker1.ReportProgress(0, "loading......");
            } while (finish != true);

        }

        void ErrorstopScript()
        {
            UseWaitCursor = false;
            btConnect.Enabled = true;
            btclear.Visible = true;
            btStop.Visible = false;
            btExport.Enabled = false;
        }

        void report(int number, string txt)
        {
            Thread.Sleep(1 / 4);
            backgroundWorker1.ReportProgress(number, txt);
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

            txtStatus.Text = e.UserState as string;
        }
        //reportprocess func
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            t1 = new Thread(new ThreadStart(loading));
            report(0, "loading.");
            string connetionString = null;
            SqlConnection cnn;

            report(0, "loading..");

            connetionString = @"Data Source=" + txtServerName.Text + ";Initial Catalog=" + txtDatabaseName.Text + "; User ID=" + txtID.Text + "; Password=" + txtPass.Text;
            cnn = new SqlConnection(connetionString);
            report(0, "loading...");
            string ResultTable = txtOutputTable.Text;
            string CalTable = txtCalTable.Text;

            string query = "SELECT DISTINCT  c.OBJECTID, c.Name AS Origin_Destetination, ODS.CustNum AS Origin_Number, ODS.CustName AS Origin_Name, ODS2.CustNum AS DesNumber, ODS2.CustName, c.Total_Kilo AS Distance_Travel, c.Total_Time AS Time_Travel " +
                "FROM((" + ResultTable + " AS c INNER JOIN " + CalTable + " AS ODS ON ODS.CustNum = c.OriginalNumber)INNER JOIN " + CalTable + " AS ODS2 ON ODS2.CustNum = c.DestinationNumber) " +
                "ORDER BY c.OBJECTID ASC;";

            report(0, "loading to table...");
            //try
            //{
            cnn.Open();

            using (SqlDataAdapter a = new SqlDataAdapter(query, cnn))
            {

                t1.Start();
                a.Fill(dt);
                dt.TableName = "ResultCalculated";
                Thread.Sleep(1);
                t1.Abort();

                report(100, "Ready to next step.");
            }
            cnn.Close();
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            if (e.Error != null)
            {
                t1.Abort();
                MessageBox.Show("Error : " + e.Error.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);             
                backgroundWorker1.CancelAsync();
                ErrorstopScript();

            }
            else if (backgroundWorker1.CancellationPending == true)
            {             
                ErrorstopScript();
            }
            else
            {
                UseWaitCursor = false;
                btclear.Visible = true;
                dataGridView1.DataSource = dt;
                btExport.Enabled = true;
                btExport.BackColor = Color.SpringGreen;
                btConnect.Enabled = false;
                txtStatus.Text = ("Process successfully");

            }
        }
        #endregion

        #region export to excel

        void loadingExport()
        {
            txtStatus.BackColor = Color.IndianRed;
            do
            {
                Thread.Sleep(500);
                backgroundWorker2.ReportProgress(0, "loading.");
                Thread.Sleep(500);
                backgroundWorker2.ReportProgress(0, "loading..");
                Thread.Sleep(500);
                backgroundWorker2.ReportProgress(0, "loading...");
                Thread.Sleep(500);
                backgroundWorker2.ReportProgress(0, "loading....");
                Thread.Sleep(500);
                backgroundWorker2.ReportProgress(0, "loading......");
            } while (finish != true);

        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            t2 = new Thread(new ThreadStart(loadingExport));
            t2.Start();

            int count = dt.Rows.Count;
            int numloop = count / 1000000;
            int loop = 0;

            if (count > 100000)
            {
                for (int j = 0; j <= numloop; j++)
                {
                    DataTable dtprocess = new DataTable();
                    foreach (DataColumn dc in dt.Columns)
                        dtprocess.Columns.Add(dc.ColumnName);
                    for (int i = 0; i < 100000; i++)
                    {
                        dtprocess.Rows.Add(dt.Rows[loop].ItemArray);
                        loop++;
                        if (loop == count) break;
                    }
                    dtprocess.TableName = "Table" + j;
                    ds.Tables.Add(dtprocess);
                }
            }
            var wb = new XLWorkbook();
            foreach (DataTable dtt in ds.Tables)
            {
                var dtEx = dtt;
                wb.Worksheets.Add(dtEx, dtt.TableName);
            }
            Console.WriteLine("start process");
            wb.SaveAs(@path);
            Console.WriteLine("finish");
            MessageBox.Show("Export successfully.");
            t2.Abort();
        }

        private void backgroundWorker2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            txtStatus.Text = e.UserState as string;
        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            UseWaitCursor = false;
            if (e.Error != null)
            {
                btConnect.Enabled = false;
                MessageBox.Show("Error :" + e.Error.ToString());
                UseWaitCursor = false;
            }
            else if (backgroundWorker2.CancellationPending == true)
            {
                UseWaitCursor = false;
                t2.Abort();

            }
            else
            {
                Cursor.Current = Cursors.Default;
                Process.Start(path);
                btclear.Enabled = true;
            }
        }
        #endregion

        //connectButton
        private void btConnect_Click(object sender, EventArgs e)
        {
            UseWaitCursor = true;
            btConnect.Enabled = false;
            if (checkBox1.Checked)
            {
              
            }
            backgroundWorker1.RunWorkerAsync();
            btclear.Visible = false;
            btStop.Visible = true;
        }
        //exportButton
        private void btExport_Click(object sender, EventArgs e)
        {
            fileName = null;
            savePath = null;
            path = null;
            sfd.Filter = ".xlsx Files (*.xlsx)|*.xlsx";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                fileName = Path.GetFileName(sfd.FileName);
                savePath = sfd.FileName.Replace(@"\" + fileName, "");
                path = sfd.FileName;
                txtStatus.Text = "Creating xlsx file at:" + path;
            }

            if (path != null)
            {
                backgroundWorker2.RunWorkerAsync();
                UseWaitCursor = true;
                btclear.Enabled = false;
            }
            else
                MessageBox.Show("Error : Please sele location to save.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        //clearButton
        private void btclear_Click(object sender, EventArgs e)
        {
            dt.Clear();
            ds.Clear();
            if (dataGridView1.Columns.Count != 0) dataGridView1.Columns.Clear();
            if (dt != null) dt = new DataTable();
            if (ds != null) ds = new DataSet();


            // if (dataGridView1.Columns.Count != 0) dataGridView1.Columns.Clear();          
            btConnect.Enabled = true;
            btExport.Enabled = false;
            txtStatus.Text = "Standby...";
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            //connectDB
            if (e.KeyCode==Keys.C&& (ModifierKeys & Keys.Control) == Keys.Control)          
                btConnect.PerformClick();
            else
                Console.WriteLine("test " + e.KeyCode);

            //Export
            if (e.KeyCode == Keys.S && (ModifierKeys & Keys.Control) == Keys.Control)
                btExport.PerformClick();
            else
                Console.WriteLine("test " + e.KeyCode);

            //clear
            if (e.KeyCode == Keys.D && (ModifierKeys & Keys.Control) == Keys.Control)
                btclear.PerformClick();
            else
                Console.WriteLine("test " + e.KeyCode);
        }
        //stopBackgroundWorker
        private void btStop_Click(object sender, EventArgs e)
        {
        
            backgroundWorker1.CancelAsync();
        }


    }
}
