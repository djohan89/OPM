﻿using OPM.OPMEnginee;
using OPM.WordHandler;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace OPM.GUI
{
    public partial class NTKTInfor : Form
    {
        public delegate void UpdateCatalogDelegate(string value);
        public UpdateCatalogDelegate UpdateCatalogPanel;

        public delegate void RequestDashBoardOpenNTKTForm(string strIDContract, string strKHMS, string strPONumber, string strPOID);
        public RequestDashBoardOpenNTKTForm requestDashBoardOpenNTKTForm;

        public delegate void RequestDashBoardPurchaseOderForm(string strIDPO, string KHMS);
        public RequestDashBoardPurchaseOderForm requestDashBoardPurchaseOderForm;

        

        public NTKTInfor()
        {
            InitializeComponent();
        }

        public void SetKHMS(string value)
        {
            txbKHMS.Text = value;
            return;
        }
        public void SetContractID(string value)
        {
            txbIDContract.Text = value;
            return;
        }
        public void SetPOID(string value)
        {
            txbPOID.Text = value;
            return;
        }
        public void SetPONumber(string value)
        {
            txbPONumber.Text = value;
            return;
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            NTKT newNKTTObj = new NTKT();
            newNKTTObj.KHMS = txbKHMS.Text;
            newNKTTObj.IDContract = txbIDContract.Text;
            newNKTTObj.POID = txbPOID.Text;
            newNKTTObj.PONumber = txbPONumber.Text;
            newNKTTObj.ID_NTKT = txbNTKTID.Text;
            newNKTTObj.DateDuKienNTKT = dateTimePickerNTKT.Value.ToString("yyyy-MM-dd");
            newNKTTObj.MrPhoBan = txbForBan.Text;
            newNKTTObj.MrPhoBanMobile = txbMobileForBan.Text;
            newNKTTObj.MrGD_CSKH = txbGDCSKH.Text;
            newNKTTObj.MrGD_CSKH_mobile = txbMobileGDCSKH.Text;
            newNKTTObj.MrGD_CSKH_Landline = txbLandLineGDCSKH.Text;
            newNKTTObj.MrrGD_CSKH_LandlineExt = txbExt.Text;
            int nod = Convert.ToInt32(txbNoD.Text);
            newNKTTObj.NumberOfDevice = (float)nod;

            int ret = 0;
            /*Create Folder NTKT*/
            //string strContractDirectory = txbIDContract.Text.Replace('/', '_');
            //strContractDirectory = strContractDirectory.Replace('-', '_');
            //string strPODirectory = "F:\\OPM\\" + strContractDirectory + "\\" + txbPONumber.Text + "\\" + "NTKT_" +txbNTKTID.Text.ToString();
            string DriveName = "";
            DriveInfo[] driveInfos = DriveInfo.GetDrives();
            foreach (DriveInfo driveInfo in driveInfos)
            {
                //MessageBox.Show(driveInfo.Name.ToString());
                if (String.Compare(driveInfo.Name.ToString().Substring(0, 3), @"D:\") == 0 || String.Compare(driveInfo.Name.ToString().Substring(0, 3), @"E:\") == 0)
                {
                    //MessageBox.Show(driveInfo.Name.ToString().Substring(0, 1));
                    DriveName = driveInfo.Name.ToString().Substring(0, 3);
                    break;
                }
            }
            Directory.CreateDirectory(DriveName + "OPM");
            Directory.CreateDirectory(DriveName + "OPM" + txbPONumber.Text);
            Directory.CreateDirectory(DriveName + "OPM" + txbPONumber.Text + "\\" + "NTKT_" + txbNTKTID.Text.ToString());
            string strPODirectory = DriveName + "OPM\\" + txbPONumber.Text + "\\" + "NTKT_" + txbNTKTID.Text.ToString();
            ret = newNKTTObj.CheckExistNTKT(txbNTKTID.Text);
            if (0 == ret)
            {
                if (!Directory.Exists(strPODirectory))
                {
                    Directory.CreateDirectory(strPODirectory);
                    MessageBox.Show("Folder " + txbPONumber.Text+ " have been created!!!");
                }

                else
                {
                    MessageBox.Show("Folder "+ txbPONumber.Text + " already exist!!!");

                }
                ret = newNKTTObj.InsertNewNTKT(newNKTTObj);
                if (0 == ret)
                {
                    MessageBox.Show(ConstantVar.CreateNewNTKTFail);
                }
                else
                {
                    MessageBox.Show(ConstantVar.CreateNewNTKTSuccess);
                    UpdateCatalogPanel("NTKT_"+txbNTKTID.Text);
                    /*Create Bao Lanh Thuc Hien Hop Dong*/
                    Directory.CreateDirectory(DriveName + "LP");
                    //string fileRQNTKT_temp = @"F:\LP\NTKT_Request_template.docx";
                    //string fileSofware_Certificate_Template = @"F:\LP\Sofware_Certificate_Template.docx";
                    //string fileCNCL = @"F:\LP\GIAY CHUNG NHAN CHAT LUONG_TONG_HOP_Template.docx";
                    string fileRQNTKT_temp = DriveName + @"LP\NTKT_Request_template.docx";
                    string fileSofware_Certificate_Template = DriveName + @"\LP\Sofware_Certificate_Template.docx";
                    string fileCNCL = DriveName + @"LP\GIAY CHUNG NHAN CHAT LUONG_TONG_HOP_Template.docx";
                    string strRQNTKTName = strPODirectory + "\\CV De Nghi NTKT_" + txbPONumber.Text +"_"+ txbIDContract.Text +".docx";
                    string strSofware_Certificate = strPODirectory + "\\Chung chi ban quyen phan mem" + txbPONumber.Text + "_" + txbIDContract.Text + ".docx";
                    string strCNCL = strPODirectory + "\\Chung nhan chat luong" + txbPONumber.Text + "_" + txbIDContract.Text + ".docx";
                    strRQNTKTName = strRQNTKTName.Replace("/","_");
                    strSofware_Certificate = strSofware_Certificate.Replace("/", "_");
                    strCNCL= strCNCL.Replace("/", "_");
                    ContractObj contractObj= new ContractObj();
                    int ret0 = ContractObj.GetObjectContract(txbIDContract.Text, ref contractObj);
                    PO pO = new PO();
                    int ret1 = PO.GetObjectPO(txbPOID.Text, ref pO);
                    this.Cursor = Cursors.WaitCursor;
                    NTKT nTKT = new NTKT();
                    nTKT.GetObjectNTKTByIDPO(txbPOID.Text, ref nTKT);
                    int num = Convert.ToInt32(txbNoD.Text);
                    nTKT.NumberOfDevice = (float)num;
                    float numofD = (float)num;
                    OpmWordHandler.Create_RQNTKT_PO(fileRQNTKT_temp, strRQNTKTName, nTKT, pO, contractObj);
                    OpmWordHandler.Create_Sofware_Certificate_Template(fileSofware_Certificate_Template, strSofware_Certificate, contractObj.IdContract, txbPONumber.Text,txbKHMS.Text, txbNTKTID.Text, numofD);
                    OpmWordHandler.Create_CNCL(fileCNCL, strCNCL, contractObj.IdContract, txbPONumber.Text, txbKHMS.Text, txbNTKTID.Text, numofD);
                    /////////////////////////////
                    
                    

                    /*Create Bao Lanh Thuc Hien Hop Dong*/
                    int ret2 = 0;
                    string fileBBKTKTHH_temp = @"F:\LP\Bien_Ban_KTKT_HH_Template.docx";
                    string strBBKTKT = strPODirectory + "\\Biên Bản Kiểm Tra Kỹ Thuật_" + txbPONumber.Text + "_" + txbIDContract.Text + ".docx";
                    strBBKTKT = strBBKTKT.Replace("/", "_");
                    
                    SiteInfo siteInfoB = new SiteInfo();
                    SiteInfo siteInfoA = new SiteInfo();
                    siteInfoB.GetSiteInfoObject(txbIDContract.Text, ref siteInfoB);
                    siteInfoA.GetSiteInfoA(txbIDContract.Text, ref siteInfoA);
                    this.Cursor = Cursors.WaitCursor;
                    OpmWordHandler.Create_BBKTKT_HH(fileBBKTKTHH_temp, strBBKTKT, contractObj, pO, nTKT, siteInfoB, siteInfoA);
                    this.Cursor = Cursors.Default;
                    
                }
                /*Create File Nghiệm Thu Kỹ Thuật*/
                /*Request Update Catalog Admin*/
            }
        }
        private void btnBack_Click(object sender, EventArgs e)
        {
            if(txbPOID.Text != null)
            {
                //ContractInfoChildForm contractInfoChildForm = new ContractInfoChildForm();
                //contractInfoChildForm.RequestDashBoardOpenPOForm = new ContractInfoChildForm.RequestDashBoardOpenChildForm(OP)
                requestDashBoardPurchaseOderForm(txbPOID.Text, txbKHMS.Text);
                PurchaseOderInfor purchaseOderInfor = new PurchaseOderInfor();
                //purchaseOderInfor.SetValueItemForPO();

            }
            else
            {
                //tra ve form rong
            }
            return;
        }
        //@Dưỡng Bùi -- Show thông tin NTKT lên UI
        public void setValueItemForNTKT(string IDNTKT)
        {
            NTKT nTKT = new NTKT();
            nTKT.GetObjectNTKT(IDNTKT, ref nTKT);
            string idPO = null, poNumber = null, idContract = null;
            int ret = nTKT.getPOinfor(IDNTKT, ref idPO, ref poNumber, ref idContract);
            PO pO = new PO();
            string namecontract = null, KHMS = null;
            pO.DisplayPO(idPO, ref namecontract, ref KHMS);
            this.txbKHMS.Text = (string)KHMS;
            this.txbIDContract.Text = (string)idContract;
            this.txbPOID.Text = (string)nTKT.POID;
            this.txbPONumber.Text = (string)poNumber;
            this.txbNTKTID.Text = (string)nTKT.ID_NTKT;
            dateTimePickerNTKT.Value = Convert.ToDateTime(nTKT.DateDuKienNTKT);
        }

        private void NTKTInfor_Load(object sender, EventArgs e)
        {

        }
    }
}
