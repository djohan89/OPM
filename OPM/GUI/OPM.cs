﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using OPM.EmailHandler;
using OPM.ExcelHandler;
using OPM.DBHandler;
using OPM.OPMEnginee;
namespace OPM
{
    public partial class OPM : Form
    {
        public OPM()
        {
            InitializeComponent();

            string filename = ConstantVar.DocTemplateDir+"MST.docx";


            OPMEmailHandler oPMEmailHandler = new OPMEmailHandler();
            //OPMEmailHandler.fSendEmail("AAAA");
            //See if any printers are installed  
            //if (PrinterSettings.InstalledPrinters.Count <= 0)
            //{
            //    MessageBox.Show("Printer not found!");
            //    return;
            //}

            ////Get all available printers and add them to the combo box  
            //List<string> printersList = new List<string>();
            //foreach (String printer in PrinterSettings.InstalledPrinters)
            //{
            //    printersList.Add(printer.ToString());
            //}
        }



        private void twPanel_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                //TreeNode selectedNode = twPanel.GetNodeAt(e.X, e.Y);
                //MessageBox.Show("You clicked on node: " + selectedNode.Text);
                /*Check Type of Note*/
                /*Display the Context Menu*/
                ctmPanel.Show();
                

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {

            //OpmWordHandler _wordhandle = new OpmWordHandler();
            //_wordhandle.FileName="DoanTD";
            //_wordhandle.fCreate_BLTU_Contract(@"F:\07. Digital Tranfrom\Tài liệu quy trình\Template\Contract\Bảo Lãnh Tạm Ứng.docx", @"F:\07. Digital Tranfrom\Tài liệu quy trình\OPM_Doc\Contract\Bảo Lãnh Tạm Ứng.docx", "PO1", "SSDCDCMS-050121-1113-19", "29/10/1989", "22/01/1989");
            //OpmExcelHandler _excelhandler = new OpmExcelHandler();
            //_excelhandler.FileName = "";
            //LogHelper.Writelog("AAAAAAAAAAAAAAAAAAAAAAAAAAAA");

            //string Filepath = @"E:\Cetificate\AAA.pdf";
            //string Filename = "AAA.pdf";
            //string PrinterName = "PCL6 V4 Driver for Universal Print";
            //IPrinter printer = new Printer();
            //printer.PrintRawFile(PrinterName, Filepath, Filename);

            /*OK for Printer*/
            //OpmWordHandler _opmWordHandler = new OpmWordHandler();
            //_opmWordHandler.fPrintDocument(@"E:\Cetificate\AAA.doc");

            /*OK for Sending Email*/
            //OPMEmailHandler.fSendEmail("AAAAA");

            //String strconnection = @"Data Source=JOHAN-LAPTOP\SQLOPM; Initial Catalog = OpmDB; User ID = sa; Password=Pa$$w0rd";
            //String strconnection = @"Server=10.2.8.96,1433; Database = OpmDB; User ID = sa; Password=Pa$$w0rd";
            //String strconnection = @"Data Source = 127.0.0.1,1433; Initial Catalog = OpmDB; User ID = sa; Password = Pa$$w0rd";
            
            //String OK_1
            //String strconnection1 = @"Data Source=DOANTD; Initial Catalog = OpmDB; User ID = sa; Password=Pa$$w0rd";

            //String strconnection2 = @"Data Source=MSI\PHANTOM_OPM; Initial Catalog = OpmDB; User ID = sa; Password=Pa$$w0rd";
            //SqlConnection con = new SqlConnection(strconnection2);
            //con.Open();

            string querry = @"INSERT INTO Site_Info(id, type, headquater_info, address, phonenumber, tin, account, representative) VALUES('TTCUVT-TPDN', '', '125 Hai Ba Trung, TP.HCM', '12/1 Nguyen Thi Minh Khai, TP.HCM', '02835282338', '0300954529', '0071001103921', 'Mr Ho Minh Kiet')";
            //using (SqlCommand insertCommand = con.CreateCommand())
            //{
            //    insertCommand.CommandText = querry;
            //    var row = insertCommand.ExecuteNonQuery();
            //}    
            //con.Close();
            //con.Dispose();
            //con = null;
            OPMDBHandler.fInsertData(querry);


         }

        private void button7_Click(object sender, EventArgs e)
        {
            DataSet result = new DataSet();
            
            string[] Chosen_Files=null;
            try
            {
                List<Packagelist> oPackagelists = new List<Packagelist>();
                //string[] Chosen_File = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                //openFileDialog.InitialDirectory = Chosen_File;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    Chosen_Files = openFileDialog.FileNames;
                }
                if (Chosen_Files.Count() == 0)
                {
                    return;
                }
                //tbx_plfile.Text = Chosen_File;

                int ret = OpmExcelHandler.fReadPackageListFiles(Chosen_Files, ref oPackagelists);
                /**/
                /**/
            }
            catch (Exception)
            {
                MessageBox.Show("Sorry, Error");
                return;

            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                string Chosen_File = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                openFileDialog.InitialDirectory = Chosen_File;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    Chosen_File = openFileDialog.FileName;
                }
                if (Chosen_File.ToString() == string.Empty)
                {
                    return;
                }

                int ret = OpmExcelHandler.fReadExcelFile(Chosen_File);
                if (0 == ret)
                {
                    return;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Sorry, Error");
                return;

            }

        }

        private void button5_Click(object sender, EventArgs e)
        {

        }
        
        private void OPM_Load(object sender, EventArgs e)
        {
            //GUI.Contract_Info contract_InfoDashboard = new GUI.Contract_Info();
            //contract_InfoDashboard.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            //| System.Windows.Forms.AnchorStyles.Left) 
            //| System.Windows.Forms.AnchorStyles.Right)));
            //contract_InfoDashboard.Location = new System.Drawing.Point(20, 70);
            //contract_InfoDashboard.Name = "contract_InfoDashboard";
            //contract_InfoDashboard.Size = new System.Drawing.Size(562, 518);
            //contract_InfoDashboard.TabIndex = 0;

            //this.splitContainer1.Panel2.Controls.Add(contract_InfoDashboard);
            ////contract_InfoDashboard.Hide();
            ///

            GUI.UsrCtltabHeader usrCtltabHeader = new GUI.UsrCtltabHeader();
            this.splitContainer1.Panel1.Controls.Add(usrCtltabHeader);
            usrCtltabHeader.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
        }
    }
}

