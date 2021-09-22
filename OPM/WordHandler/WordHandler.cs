﻿using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;
using WordOffice = Microsoft.Office.Interop.Word;
using System.Reflection;
using OPM.OPMEnginee;
using OPM.GUI;
namespace OPM.WordHandler
{
    class OpmWordHandler
    {
        private string _nameWordfile;
        public OpmWordHandler()
        {
        }
        ~OpmWordHandler()
        { }

        public string FileName
        {
            set { _nameWordfile = value; }
            get { return _nameWordfile; }
        }
        public static void FindAndReplace(WordOffice.Application wordApp, object ToFindText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref ToFindText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundLike,
                ref nmatchAllforms, ref forward,
                ref wrap, ref format, ref replaceWithText,
                ref replace, ref matchKashida,
                ref matchDiactitics, ref matchAlefHamza,
                ref matchControl);
        }
        private static ConvertDateFormat convertDateFormat = new ConvertDateFormat();
        public static void CreateWordDocument(object filename, object SaveAs, string strName, string strFirstname, string strBirthday, string strDate)
        {
            WordOffice.Application wordApp = new WordOffice.Application();
            object missing = Missing.Value;
            WordOffice.Document myWordDoc = null;

            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing, ref missing);
                myWordDoc.Activate();

                //find and replace
                FindAndReplace(wordApp, "<name>", strName);
                FindAndReplace(wordApp, "<firstname>", strFirstname);
                //this.FindAndReplace(wordApp, "<birthday>", dateTimePicker1.Value.ToShortDateString());
                FindAndReplace(wordApp, "<birthday>", strBirthday);
                //this.FindAndReplace(wordApp, "<date>", DateTime.Now.ToShortDateString());
                string[] dates = convertDateFormat.ConvertFormatDate(strDate, "yyyy-MM-dd", "dd/MM/yyyy");
                FindAndReplace(wordApp, "<date>", dates[0]+"/"+dates[1]+"/"+dates[2]+" ");
            }
            else
            {
                MessageBox.Show("File not Found!");
            }

            //Save as
            myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing);

            myWordDoc.Close();
            wordApp.Quit();
            MessageBox.Show("File Created!");
        }

        public static void Create_BLTU_Contract(object filename, object SaveAs, string strPOnumber, string strIdContract, string strSigndate, string strPOdate)
        {
            //thư viện Microsoft.Office.Interop.Word
            WordOffice.Application wordApp = new WordOffice.Application();
            object missing = Missing.Value;
            WordOffice.Document myWordDoc = null;

            try
            {
                if (File.Exists((string)filename))
                {
                    object readOnly = false;
                    object isVisible = false;
                    wordApp.Visible = false;

                    myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing, ref missing);
                    myWordDoc.Activate();

                    //find and replace
                    FindAndReplace(wordApp, "<tempPO>", " " + strPOnumber + " ");
                    FindAndReplace(wordApp, "<temp_IdContract>", " " + strIdContract + " ");
                    //this.FindAndReplace(wordApp, "<birthday>", dateTimePicker1.Value.ToShortDateString());
                    FindAndReplace(wordApp, "<tmp_signdate>", " " + strSigndate + " ");
                    //this.FindAndReplace(wordApp, "<date>", DateTime.Now.ToShortDateString());
                    string[] date_po = convertDateFormat.ConvertFormatDate(strPOdate, "yyyy-MM-dd", "dd/MM/yyyy");
                    FindAndReplace(wordApp, "<date_po>", " " + date_po[0] + "/" + date_po[1] + "/" + date_po[2] + " ");
                }
                else
                {
                    MessageBox.Show("File not Found!");
                }

                //Save as
                myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing);

                myWordDoc.Close();
                wordApp.Quit();
                MessageBox.Show("File Created!");
            }
            catch(Exception)
            {
                myWordDoc.Close();
                wordApp.Quit();
                MessageBox.Show("File did'nt Created!");
            }
        }
        public static void Create_BLTH_PO(object filename, object SaveAs,string strPOnumber)
        { }
        public static void Create_BLTU_PO(object filename, object SaveAs,string strPOnumber, string strIdContract, string strContractName, string strSigneddateContract, string strPOSigneddate, string strPOValueNotVAT, string strPOValueTU, string strSiteB, string strActiveDatePO)
        {
            WordOffice.Application wordApp = new WordOffice.Application();
            object missing = Missing.Value;
            WordOffice.Document myWordDoc = null;

            try
            {
                if (File.Exists((string)filename))
                {
                    object readOnly = false;
                    object isVisible = false;
                    wordApp.Visible = false;
                    //Khởi tạo đói tượng myWordDoc để điền thông tin vào Word
                    myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                              ref missing, ref missing, ref missing,
                                              ref missing, ref missing, ref missing,
                                              ref missing, ref missing, ref missing,
                                              ref missing, ref missing, ref missing, ref missing);
                    //Bắt đầu thực hiện chạy chương trình đọc thông tin 
                    myWordDoc.Activate();

                    //find and replace
                    FindAndReplace(wordApp, "<PO_Name>", " " + strPOnumber);
                    FindAndReplace(wordApp, "<Contract_ID>", " " + strIdContract);
                    FindAndReplace(wordApp, "<Contract_Name>", " " + strContractName);
                    string[] Signed_DateContract = convertDateFormat.ConvertFormatDate(strSigneddateContract, "dd-MM-yyyy", "dd/MM/yyyy");
                    FindAndReplace(wordApp, "<Signed_DateContract>", " " + Signed_DateContract[0] + "/" + Signed_DateContract[1] + "/" + Signed_DateContract[2]);
                    string[] Signed_DatePO = convertDateFormat.ConvertFormatDate(strPOSigneddate, "yyyy-MM-dd", "dd/MM/yyyy");
                    FindAndReplace(wordApp, "<Signed_DatePO>", " " + Signed_DateContract[0] + "/" + Signed_DateContract[1] + "/" + Signed_DateContract[2]);
                    FindAndReplace(wordApp, "<Total_Value>", " " + strPOValueNotVAT);
                    FindAndReplace(wordApp, "<Value_Tamung>", " " + strPOValueTU);
                    FindAndReplace(wordApp, "<Site_B>", " " + strSiteB);
                    string[] Active_Date = convertDateFormat.ConvertFormatDate(strActiveDatePO, "yyyy-MM-dd", "dd/MM/yyyy");
                    FindAndReplace(wordApp, "<Active_Date>", " " + Active_Date[0] + "/" + Active_Date[1] + "/" + Active_Date[2]);
                }
                else
                {
                    MessageBox.Show("File not Found!");
                }

                //Save as
                myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing);

                myWordDoc.Close();
                wordApp.Quit();
                MessageBox.Show("File Bảo Lãnh Thực Hiện Created!");
            }
            catch
            {
                myWordDoc.Close();
                wordApp.Quit();
                MessageBox.Show("File Bảo Lãnh Thực Hiện không được Created!");
            }

        }

        public static void Create_DNTU_PO(object filename, object SaveAs)
        { }
        public static void Create_RQNTKT_PO(object filename, object SaveAs, NTKT nTKT, PO objPO, ContractObj contractObj)
        {
            WordOffice.Application wordApp = new WordOffice.Application();
            object missing = Missing.Value;
            WordOffice.Document myWordDoc = null;

            try
            {
                if (File.Exists((string)filename))
                {
                    object readOnly = false;
                    object isVisible = false;
                    wordApp.Visible = false;

                    myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing, ref missing);
                    myWordDoc.Activate();

                    //find and replace
                    FindAndReplace(wordApp, "<dd>", " " + DateTime.Now.ToString("dd") + " ");
                    FindAndReplace(wordApp, "<MM>", " " + DateTime.Now.ToString("MM") + " ");
                    FindAndReplace(wordApp, "<yyyy>", " " + DateTime.Now.ToString("yyyy") + " ");

                    FindAndReplace(wordApp, "<NTKT_ID>", " " + nTKT.ID_NTKT + " ");
                    string[] Date_NTKT_DuKien = convertDateFormat.ConvertFormatDate(nTKT.DateDuKienNTKT, "yyyy-MM-dd", "dd/MM/yyyy");
                    FindAndReplace(wordApp, "<Date_NTKT_DuKien>", " " + Date_NTKT_DuKien[0] + "/" + Date_NTKT_DuKien[1] + "/" + Date_NTKT_DuKien[2]);

                    FindAndReplace(wordApp, "<PO_Number>", " " + nTKT.PONumber + " ");
                    FindAndReplace(wordApp, "<PO_ID>", " " + nTKT.POID + " ");
                    string[] Created_DatePO = convertDateFormat.ConvertFormatDate(objPO.DateCreatedPO, "yyyy-MM-dd", "dd/MM/yyyy");
                    FindAndReplace(wordApp, "<Created_DatePO>", " " + Created_DatePO[0] + "/" + Created_DatePO[1] + "/" + Created_DatePO[2]);

                    FindAndReplace(wordApp, "<KHMS>", " " + nTKT.KHMS + " ");
                    FindAndReplace(wordApp, "<Contract_ID>", " " + nTKT.IDContract + " ");
                    FindAndReplace(wordApp, "<Contract_Name>", " " + contractObj.NameContract + " ");
                    string[] SignedDate_Contract = convertDateFormat.ConvertFormatDate(contractObj.DateSigned, "dd-MM-yyyy", "dd/MM/yyyy");
                    FindAndReplace(wordApp, "<SignedDate_Contract>", " " + SignedDate_Contract[0] + "/" + SignedDate_Contract[1] + "/" + SignedDate_Contract[2]);
                    FindAndReplace(wordApp, "<Site_B>", " " + contractObj.SiteB + " ");

                    FindAndReplace(wordApp, "<Mr_PhoBan>", " " + nTKT.MrPhoBan + " ");
                    FindAndReplace(wordApp, "<Mobile>", " " + nTKT.MrPhoBanMobile + " ");

                    FindAndReplace(wordApp, "<Mr_GD_HTKT_CSKH>", " " + nTKT.MrGD_CSKH + " ");
                    FindAndReplace(wordApp, "<Mr_GD_Mobile>", " " + nTKT.MrGD_CSKH_mobile + " ");

                    FindAndReplace(wordApp, "<MrGDLandLine>", " " + nTKT.MrGD_CSKH_Landline + " ");
                    FindAndReplace(wordApp, "<Ext>", " " + nTKT.MrrGD_CSKH_LandlineExt + " ");

                }
                else
                {
                    MessageBox.Show("File not Found!");
                }

                //Save as
                myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing);

                myWordDoc.Close();

                wordApp.Quit();
                MessageBox.Show("File Yêu Cầu Nghiệm Thu Kỹ Thuật Đã Được Created!");
            }
            catch
            {
                myWordDoc.Close();

                wordApp.Quit();
                MessageBox.Show("File Yêu Cầu Nghiệm Thu Kỹ Thuật không Được Created thành công!");
            }
        }

        public static void PrintDocument(object filename)
        {
            WordOffice.Application wordApp = new WordOffice.Application();
            object missing = Missing.Value;
            WordOffice.Document myWordDoc = null;

            try
            {
                if (File.Exists((string)filename))
                {
                    object readOnly = false;
                    object isVisible = false;
                    wordApp.Visible = false;

                    myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing, ref missing);
                    myWordDoc.Activate();

                    myWordDoc.Application.ActivePrinter = "PCL6 V4 Driver for Universal Print";
                    myWordDoc.PrintOut();
                }
                else
                {
                    MessageBox.Show("File not Found!");
                }

                myWordDoc.Close();
                wordApp.Quit();
                MessageBox.Show("File Created!");
            }
            catch(Exception)
            {
                myWordDoc.Close();
                wordApp.Quit();
                MessageBox.Show("File Created không thành công!");

            }
        }

        public static void PrintDocumentA(object filename)
        {
            return;
        }

        public static void  CreateBLTH_Contract(object filename, object SaveAs, string strContractCode, string strContractName, string strSigneddate,string tbxSiteB,string txbGaranteeValue,string txbGaranteeActiveDate)
        {
            WordOffice.Application wordApp = new WordOffice.Application();
            object missing = Missing.Value;
            WordOffice.Document myWordDoc = null;

            try
            {
                if (File.Exists((string)filename))
                {
                    object readOnly = false;
                    object isVisible = false;
                    wordApp.Visible = false;

                    myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing, ref missing);
                    myWordDoc.Activate();

                    //find and replace
                    FindAndReplace(wordApp, "<Contract_Code>", " " + strContractCode + " ");
                    FindAndReplace(wordApp, "<Contract_Name>", " " + strContractName + " ");
                    FindAndReplace(wordApp, "<Signed_Date>", " " + strSigneddate + " ");
                    FindAndReplace(wordApp, "<Site_B>", " " + tbxSiteB + " ");
                    FindAndReplace(wordApp, "<Grt_Val>", " " + txbGaranteeValue + " ");
                    FindAndReplace(wordApp, "<Grt_Act_Date>", " " + txbGaranteeActiveDate + " ");

                }
                else
                {
                    MessageBox.Show("File not Found!");
                }

                //Save as
                myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing);

                myWordDoc.Close();
                wordApp.Quit();
                MessageBox.Show("File Bảo Lãnh Thực Hiện Hợp Đồng Đã Được Tạo");
            }
            catch(Exception)
            {
                myWordDoc.Close();
                wordApp.Quit();
                MessageBox.Show("File Bảo Lãnh Thực Hiện Hợp Đồng không Được Tạo thành công");
            }
        }
           
        public static void Create_VBConfirm_PO(object filename, object SaveAs, ConfirmPO confirmPO, PO objPO, ContractObj contractObj)
        {
            WordOffice.Application wordApp = new WordOffice.Application();
            object missing = Missing.Value;
            WordOffice.Document myWordDoc = null;

            try
            {
                if (File.Exists((string)filename))
                {
                    object readOnly = false;
                    object isVisible = false;
                    wordApp.Visible = false;

                    myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing, ref missing);
                    myWordDoc.Activate();

                    //find and replace
                    FindAndReplace(wordApp, "<dd>", " " + DateTime.Now.ToString("dd") + " ");
                    FindAndReplace(wordApp, "<MM>", " " + DateTime.Now.ToString("MM") + " ");
                    FindAndReplace(wordApp, "<yyyy>", " " + DateTime.Now.ToString("yyyy") + " ");

                    FindAndReplace(wordApp, "<XNDH_ID>", " " + confirmPO.ConfirmPOID + " ");
                    string[] dateNow = convertDateFormat.ConvertFormatDate(DateTime.Now.ToString("dd-MM-yyyy"), "dd-MM-yyyy", "dd/MM/yyyy");
                    FindAndReplace(wordApp, "<DateNow>", " " + dateNow[0] + "/" + dateNow[1] + "/" + dateNow[2]);
                    FindAndReplace(wordApp, "<PO_Number>", " " + confirmPO.PONumber + " ");
                    FindAndReplace(wordApp, "<PO_ID>", " " + confirmPO.POID + " ");
                    string[] PO_CreatedDate = convertDateFormat.ConvertFormatDate(objPO.DateCreatedPO, "yyyy-MM-dd", "dd/MM/yyyy");
                    FindAndReplace(wordApp, "<PO_DateCreated>", " " + PO_CreatedDate[0] + "/" + PO_CreatedDate[1] + "/" + PO_CreatedDate[2]);

                    FindAndReplace(wordApp, "<KHMS>", " " + confirmPO.KHMS + " ");
                    FindAndReplace(wordApp, "<Contract_ID>", " " + confirmPO.IDContract + " ");
                    FindAndReplace(wordApp, "<Contract_Name>", " " + contractObj.NameContract + " ");
                    string[] Contract_DateSigned = convertDateFormat.ConvertFormatDate(contractObj.DateSigned, "dd-MM-yyyy", "dd/MM/yyyy");
                    FindAndReplace(wordApp, "<Contract_DateSigned>", " " + Contract_DateSigned[0] + "/" + Contract_DateSigned[1] + "/" + Contract_DateSigned[2]);
                    FindAndReplace(wordApp, "<Site_B>", " " + contractObj.SiteB + " ");

                    FindAndReplace(wordApp, "<MrPhoBan>", " " + confirmPO.MrPhoBan + " ");
                    FindAndReplace(wordApp, "<Mobile>", " " + confirmPO.MrPhoBanMobile + " ");

                    FindAndReplace(wordApp, "<MrGDCSKH>", " " + confirmPO.MrGD_CSKH + " ");
                    FindAndReplace(wordApp, "<MrGDMobile>", " " + confirmPO.MrGD_CSKH_mobile + " ");

                    FindAndReplace(wordApp, "<LandLine>", " " + confirmPO.MrGD_CSKH_Landline + " ");
                    FindAndReplace(wordApp, "<Ext>", " " + confirmPO.MrrGD_CSKH_LandlineExt + " ");

                }
                else
                {
                    MessageBox.Show("File Xác Nhận Đơn Hàng Not Found!");
                }

                //Save as
                myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing);

                myWordDoc.Close();
                wordApp.Quit();
                MessageBox.Show("File Xác Nhận Đơn Hàng Đã Được Created!");
            }
            catch(Exception)
            {
                myWordDoc.Close();
                wordApp.Quit();
                MessageBox.Show("File Xác Nhận Đơn Hàng không Được Created!");
            }
        }


        public static void Create_BBKTKT_HH(object filename, object SaveAs,ContractObj contractObject ,PO objPO, NTKT nTKT, SiteInfo siteInfo, SiteInfo siteInfoA)
        {
            WordOffice.Application wordApp = new WordOffice.Application();
            object missing = Missing.Value;
            WordOffice.Document myWordDoc = null;

            try
            {
                if (File.Exists((string)filename))
                {
                    object readOnly = false;
                    object isVisible = false;
                    wordApp.Visible = false;

                    myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing, ref missing);
                    myWordDoc.Activate();

                    //find and replace
                    FindAndReplace(wordApp, "<PO_Number>", " " + objPO.PONumber + " ");
                    FindAndReplace(wordApp, "<PO_ID>", " " + objPO.IDPO + " ");
                    string[] PO_CreatedDate = convertDateFormat.ConvertFormatDate(objPO.DateCreatedPO, "yyyy-MM-dd", "dd/MM/yyyy");
                    FindAndReplace(wordApp, "<PO_CreatedDate>", " " + PO_CreatedDate[0] + "/" + PO_CreatedDate[1] + "/" + PO_CreatedDate[2]);
                    FindAndReplace(wordApp, "<XNDH_ID>", " " + objPO.IDPO + " ");
                    string[] XNDH_CreatedDate = convertDateFormat.ConvertFormatDate(objPO.DefaultActiveDatePO, "yyyy-MM-dd", "dd/MM/yyyy");
                    FindAndReplace(wordApp, "<XNDH_CreatedDate>", " " + XNDH_CreatedDate[0] + "/" + XNDH_CreatedDate[1] + "/" + XNDH_CreatedDate[2]);
                    FindAndReplace(wordApp, "<NTKT_ID>", " " + nTKT.ID_NTKT + " ");
                    string[] NTKT_CreatedDate = convertDateFormat.ConvertFormatDate(nTKT.getCreateDate, "yyyy-MM-dd", "dd/MM/yyyy");
                    FindAndReplace(wordApp, "<NTKT_CreatedDate>", " " + NTKT_CreatedDate[0] + "/" + NTKT_CreatedDate[1] + "/" + NTKT_CreatedDate[2]);
                    FindAndReplace(wordApp, "<DC>", " " + (Math.Round(nTKT.NumberOfDevice - (nTKT.NumberOfDevice) * 0.02)) + " ");
                    FindAndReplace(wordApp, "<PDC>", " " + (Math.Round((nTKT.NumberOfDevice) * 0.02)) + " ");
                    string[] dateNow = convertDateFormat.ConvertFormatDate(DateTime.Now.ToString("dd-MM-yyyy"), "dd-MM-yyyy", "dd/MM/yyyy");
                    FindAndReplace(wordApp, "<DateNow>", " " + dateNow[0] + "/" + dateNow[1] + "/" + dateNow[2]);

                    FindAndReplace(wordApp, "<Contract_ID>", " " + contractObject.IdContract + " ");
                    FindAndReplace(wordApp, "<Contract_Name>", " " + contractObject.NameContract + " ");
                    FindAndReplace(wordApp, "<KHMS>", " " + contractObject.KHMS + " ");
                    string[] Contract_DateSigned = convertDateFormat.ConvertFormatDate(contractObject.DateSigned, "dd-MM-yyyy", "dd/MM/yyyy");
                    FindAndReplace(wordApp, "<Contract_DateSigned>", " " + Contract_DateSigned[0] + "/" + Contract_DateSigned[1] + "/" + Contract_DateSigned[2]);
                    FindAndReplace(wordApp, "<Site_B>", " " + contractObject.SiteB + " ");

                    FindAndReplace(wordApp, "<Address_Site_B>", " " + siteInfo.Address + " ");
                    FindAndReplace(wordApp, "<LandLine_Site_B>", " " + siteInfo.Phonenumber + " ");
                    FindAndReplace(wordApp, "<Fax_Site_B>", " " + siteInfo.Tin + " ");

                    FindAndReplace(wordApp, "<LandLine_Site_A>", " " + siteInfoA.Phonenumber + " ");
                    FindAndReplace(wordApp, "<Fax_Site_A>", " " + siteInfoA.Tin + " ");
                }
                else
                {
                    MessageBox.Show("File not Found!");
                }

                //Save as
                myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing);

                myWordDoc.Close();
                wordApp.Quit();
                MessageBox.Show("File Biên Bản Kiểm Tra Kĩ Thuật Hàng Hóa Được Tạo");
            }
            catch (Exception)
            {
                myWordDoc.Close();
                wordApp.Quit();
                MessageBox.Show("File Biên Bản Kiểm Tra Kĩ Thuật Hàng Hóa không Tạo thành công");
            }
        }
        // tạo chứng chỉ phần mềm
        public static void Create_Sofware_Certificate_Template(object filename, object SaveAs, string idContract, string pOnumber, string KHMS,string idNTKT, float numoD)
        {
            WordOffice.Application wordApp = new WordOffice.Application();
            object missing = Missing.Value;
            WordOffice.Document myWordDoc = null;

            try
            {
                if (File.Exists((string)filename))
                {
                    object readOnly = false;
                    object isVisible = false;
                    wordApp.Visible = false;

                    myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing, ref missing);
                    myWordDoc.Activate();
                    NTKT nTKT = new NTKT();
                    int ret = nTKT.GetObjectNTKT(idNTKT, ref nTKT);
                    //find and replace
                    FindAndReplace(wordApp, "<Contract_ID>", " " + idContract + " ");
                    FindAndReplace(wordApp, "<PO_Number>", " " + pOnumber + " ");
                    FindAndReplace(wordApp, "<KHMS>", " " + KHMS + " ");
                    FindAndReplace(wordApp, "<NoD>", " " + (numoD).ToString() + " ");
                    FindAndReplace(wordApp, "<NoD0.98>", " " + (numoD - Math.Round((numoD) * 0.02)).ToString() + " ");
                    FindAndReplace(wordApp, "<NoD0.02>", " " + Math.Round((numoD) * 0.02).ToString() + " ");
                    string[] dateNow = convertDateFormat.ConvertFormatDate(DateTime.Now.ToString("dd-MM-yyyy"), "dd-MM-yyyy", "dd/MM/yyyy");
                    FindAndReplace(wordApp, "<DateNow>", " " + dateNow[0] + "/" + dateNow[1] + "/" + dateNow[2]);
                }
                else
                {
                    MessageBox.Show("File not Found!");
                }

                //Save as
                myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing);

                myWordDoc.Close();
                wordApp.Quit();
                MessageBox.Show("File Giấy chứng nhận bản quyền phần mềm Được Tạo");
            }
            catch (Exception)
            {
                myWordDoc.Close();
                wordApp.Quit();
                MessageBox.Show("File Giấy chứng nhận bản quyền phần mềm không Tạo thành công");
            }
        }
        //tạo chứng nhận chất lượng
        public static void Create_CNCL(object filename, object SaveAs, string idContract, string pOnumber, string KHMS, string idNTKT, float Nod)
        {
            WordOffice.Application wordApp = new WordOffice.Application();
            object missing = Missing.Value;
            WordOffice.Document myWordDoc = null;

            try
            {
                if (File.Exists((string)filename))
                {
                    object readOnly = false;
                    object isVisible = false;
                    wordApp.Visible = false;

                    myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing, ref missing);
                    myWordDoc.Activate();
                    NTKT nTKT = new NTKT();
                    int ret = nTKT.GetObjectNTKT(idNTKT, ref nTKT);
                    //find and replace
                    FindAndReplace(wordApp, "<Contract_ID>", " " + idContract + " ");
                    FindAndReplace(wordApp, "<PO_Number>", " " + pOnumber + " ");
                    FindAndReplace(wordApp, "<KHMS>", " " + KHMS + " ");
                    FindAndReplace(wordApp, "<Total>", " " + (Nod).ToString() + " ");
                    FindAndReplace(wordApp, "<NofDe>", " " + (Nod - Math.Round((Nod) * 0.02)).ToString() + " ");
                    FindAndReplace(wordApp, "<BHD>", " " + Math.Round((Nod) * 0.02).ToString() + " ");
                    string[] dateNow = convertDateFormat.ConvertFormatDate(DateTime.Now.ToString("dd-MM-yyyy"), "dd-MM-yyyy", "dd/MM/yyyy");
                    FindAndReplace(wordApp, "<dd>/<MM>/<yyyy>", " " + dateNow[0] + "/" + dateNow[1] + "/" + dateNow[2]);
                }
                else
                {
                    MessageBox.Show("File not Found!");
                }

                //Save as
                myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing);

                myWordDoc.Close();
                wordApp.Quit();
                MessageBox.Show("File chứng nhận hợp quy của thiết bị được tạo");
            }
            catch (Exception)
            {
                myWordDoc.Close();
                wordApp.Quit();
                MessageBox.Show("File chứng nhận hợp quy của thiết bị không được tạo thành công");
            }
        }
    }
}
