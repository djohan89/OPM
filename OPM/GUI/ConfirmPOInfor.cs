using OPM.OPMEnginee;
using OPM.WordHandler;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
namespace OPM.GUI
{
    public partial class ConfirmPOInfor : Form
    {
        public ConfirmPOInfor()
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
            ConfirmPO newConfirmPOObj = new ConfirmPO();
            newConfirmPOObj.KHMS = txbKHMS.Text;
            newConfirmPOObj.IDContract = txbIDContract.Text;
            newConfirmPOObj.POID = txbPOID.Text;
            newConfirmPOObj.PONumber = txbPONumber.Text;
            newConfirmPOObj.ConfirmPOID = txbConfirmPOID.Text;

            newConfirmPOObj.MrPhoBan = txbForBan.Text;
            newConfirmPOObj.MrPhoBanMobile = txbMobileForBan.Text;
            newConfirmPOObj.MrGD_CSKH = txbGDCSKH.Text;
            newConfirmPOObj.MrGD_CSKH_mobile = txbMobileGDCSKH.Text;
            newConfirmPOObj.MrGD_CSKH_Landline = txbLandLineGDCSKH.Text;
            newConfirmPOObj.MrrGD_CSKH_LandlineExt = txbExt.Text;

            int ret = 0;
            /*Create Folder NTKT*/
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
            string strPODirectory = DriveName + "OPM\\" + txbPONumber.Text;
            //
            //string strContractDirectory = txbIDContract.Text.Replace('/', '_');
            //strContractDirectory = strContractDirectory.Replace('-', '_');
            //string strPODirectory = "F:\\OPM\\" + strContractDirectory + "\\" + txbPONumber.Text;
            
            ret = newConfirmPOObj.CheckExistConfirmPO(txbConfirmPOID.Text);
            if (0 == ret)
            {
                ret = newConfirmPOObj.InsertNewConfirmPO(newConfirmPOObj);
                if (0 == ret)
                {
                    MessageBox.Show(ConstantVar.CreateNewConfirmPOFail);
                }
                else
                {
                    MessageBox.Show(ConstantVar.CreateNewConfirmPOSuccess);
                    /*Create Bao Lanh Thuc Hien Hop Dong*/
                    Directory.CreateDirectory(DriveName + "LP");
                    string fileCVXNDH_temp = DriveName+@"LP\CV_XNDH_Template.docx";
                    //string fileCVXNDH_temp = @"F:\LP\CV_XNDH_Template.docx";
                    string strCVXNDHName = strPODirectory + "\\CV Xác Nhận Đơn Hàng_" + txbPONumber.Text + "_" + txbIDContract.Text + ".docx";
                    strCVXNDHName = strCVXNDHName.Replace("/", "_");
                    ContractObj contractObj = new ContractObj();
                    ret = ContractObj.GetObjectContract(txbIDContract.Text, ref contractObj);
                    PO pO = new PO();
                    ret = PO.GetObjectPO(txbPOID.Text, ref pO);
                    this.Cursor = Cursors.WaitCursor;
                    OpmWordHandler.Create_VBConfirm_PO(fileCVXNDH_temp, strCVXNDHName, newConfirmPOObj, pO, contractObj);
                    this.Cursor = Cursors.Default;
                }
                /*Create File Nghiệm Thu Kỹ Thuật*/
                /*Request Update Catalog Admin*/
            }
        }
        }
}
