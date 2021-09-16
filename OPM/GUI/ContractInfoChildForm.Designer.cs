﻿
using System.Drawing;

namespace OPM.GUI
{
    partial class ContractInfoChildForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnDescriptionA = new System.Windows.Forms.Button();
            this.tbContract = new System.Windows.Forms.TextBox();
            this.tbBidName = new System.Windows.Forms.TextBox();
            this.tbAccountingCode = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.tbxDurationContract = new System.Windows.Forms.TextBox();
            this.txbTypeContract = new System.Windows.Forms.TextBox();
            this.tbxValueContract = new System.Windows.Forms.TextBox();
            this.tbxDurationPO = new System.Windows.Forms.TextBox();
            this.tbxSiteA = new System.Windows.Forms.TextBox();
            this.tbxSiteB = new System.Windows.Forms.TextBox();
            this.btnNewPO = new System.Windows.Forms.Button();
            this.btnEdit = new System.Windows.Forms.Button();
            this.btnRemove = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCreateGarantee = new System.Windows.Forms.Button();
            this.txbGaranteeValue = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.dateTimePickerDateSignedPO = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickerDurationDateContract = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickerActiveDateContract = new System.Windows.Forms.DateTimePicker();
            this.label16 = new System.Windows.Forms.Label();
            this.txbKHMS = new System.Windows.Forms.TextBox();
            this.label17 = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.ExpirationDate = new System.Windows.Forms.DateTimePicker();
            this.txbGaranteeActiveDate = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnDescriptionA
            // 
            this.btnDescriptionA.Location = new System.Drawing.Point(527, 608);
            this.btnDescriptionA.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnDescriptionA.Name = "btnDescriptionA";
            this.btnDescriptionA.Size = new System.Drawing.Size(107, 38);
            this.btnDescriptionA.TabIndex = 0;
            this.btnDescriptionA.Text = "Desc...";
            this.btnDescriptionA.UseVisualStyleBackColor = true;
            this.btnDescriptionA.Click += new System.EventHandler(this.IdSiteA_Click);
            // 
            // tbContract
            // 
            this.tbContract.BackColor = System.Drawing.SystemColors.Window;
            this.tbContract.Location = new System.Drawing.Point(200, 58);
            this.tbContract.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tbContract.Name = "tbContract";
            this.tbContract.Size = new System.Drawing.Size(434, 31);
            this.tbContract.TabIndex = 1;
            // 
            // tbBidName
            // 
            this.tbBidName.Location = new System.Drawing.Point(200, 108);
            this.tbBidName.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tbBidName.Name = "tbBidName";
            this.tbBidName.Size = new System.Drawing.Size(434, 31);
            this.tbBidName.TabIndex = 1;
            // 
            // tbAccountingCode
            // 
            this.tbAccountingCode.Location = new System.Drawing.Point(200, 158);
            this.tbAccountingCode.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tbAccountingCode.Name = "tbAccountingCode";
            this.tbAccountingCode.Size = new System.Drawing.Size(434, 31);
            this.tbAccountingCode.TabIndex = 1;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(527, 658);
            this.button2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(107, 38);
            this.button2.TabIndex = 2;
            this.button2.Text = "Desc...";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.IdSiteB_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Location = new System.Drawing.Point(14, 67);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(127, 25);
            this.label1.TabIndex = 3;
            this.label1.Text = "Mã Hợp Đồng";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(14, 117);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(113, 25);
            this.label2.TabIndex = 3;
            this.label2.Text = "Tên Gói Thầu";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(14, 167);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(103, 25);
            this.label3.TabIndex = 3;
            this.label3.Text = "Mã Kế Toán";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(14, 217);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(77, 25);
            this.label4.TabIndex = 3;
            this.label4.Text = "Ngày Ký";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(14, 267);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(167, 25);
            this.label5.TabIndex = 3;
            this.label5.Text = "Thời Hạn Thực Hiện";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(14, 317);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(134, 25);
            this.label6.TabIndex = 3;
            this.label6.Text = "Loại Hợp Đồng";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(14, 367);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(127, 25);
            this.label7.TabIndex = 3;
            this.label7.Text = "Ngày Hiệu Lực";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(14, 417);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(149, 25);
            this.label8.TabIndex = 3;
            this.label8.Text = "Giá Trị Hợp Đồng";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(14, 467);
            this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(160, 25);
            this.label9.TabIndex = 3;
            this.label9.Text = "Thời Hạn Bảo Lãnh";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(14, 617);
            this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(53, 25);
            this.label10.TabIndex = 3;
            this.label10.Text = "SiteA";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(14, 667);
            this.label11.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(51, 25);
            this.label11.TabIndex = 3;
            this.label11.Text = "SiteB";
            // 
            // tbxDurationContract
            // 
            this.tbxDurationContract.Location = new System.Drawing.Point(200, 258);
            this.tbxDurationContract.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tbxDurationContract.Name = "tbxDurationContract";
            this.tbxDurationContract.Size = new System.Drawing.Size(110, 31);
            this.tbxDurationContract.TabIndex = 1;
            this.tbxDurationContract.Text = "0";
            this.tbxDurationContract.TextChanged += new System.EventHandler(this.tbxDurationContract_TextChanged);
            // 
            // txbTypeContract
            // 
            this.txbTypeContract.Location = new System.Drawing.Point(200, 308);
            this.txbTypeContract.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txbTypeContract.Name = "txbTypeContract";
            this.txbTypeContract.Size = new System.Drawing.Size(434, 31);
            this.txbTypeContract.TabIndex = 1;
            // 
            // tbxValueContract
            // 
            this.tbxValueContract.Location = new System.Drawing.Point(200, 408);
            this.tbxValueContract.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tbxValueContract.Name = "tbxValueContract";
            this.tbxValueContract.Size = new System.Drawing.Size(355, 31);
            this.tbxValueContract.TabIndex = 1;
            this.tbxValueContract.Text = "0";
            this.tbxValueContract.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // tbxDurationPO
            // 
            this.tbxDurationPO.Location = new System.Drawing.Point(200, 458);
            this.tbxDurationPO.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tbxDurationPO.Name = "tbxDurationPO";
            this.tbxDurationPO.Size = new System.Drawing.Size(355, 31);
            this.tbxDurationPO.TabIndex = 1;
            this.tbxDurationPO.Text = "0";
            this.tbxDurationPO.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // tbxSiteA
            // 
            this.tbxSiteA.Location = new System.Drawing.Point(200, 608);
            this.tbxSiteA.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tbxSiteA.Name = "tbxSiteA";
            this.tbxSiteA.Size = new System.Drawing.Size(287, 31);
            this.tbxSiteA.TabIndex = 1;
            // 
            // tbxSiteB
            // 
            this.tbxSiteB.Location = new System.Drawing.Point(200, 658);
            this.tbxSiteB.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tbxSiteB.Name = "tbxSiteB";
            this.tbxSiteB.Size = new System.Drawing.Size(287, 31);
            this.tbxSiteB.TabIndex = 1;
            // 
            // btnNewPO
            // 
            this.btnNewPO.Location = new System.Drawing.Point(527, 762);
            this.btnNewPO.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnNewPO.Name = "btnNewPO";
            this.btnNewPO.Size = new System.Drawing.Size(107, 38);
            this.btnNewPO.TabIndex = 4;
            this.btnNewPO.Text = "New PO";
            this.btnNewPO.UseVisualStyleBackColor = true;
            this.btnNewPO.Click += new System.EventHandler(this.btnNewPO_Click);
            // 
            // btnEdit
            // 
            this.btnEdit.Location = new System.Drawing.Point(290, 762);
            this.btnEdit.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(107, 38);
            this.btnEdit.TabIndex = 4;
            this.btnEdit.Text = "Edit";
            this.btnEdit.UseVisualStyleBackColor = true;
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnRemove
            // 
            this.btnRemove.Location = new System.Drawing.Point(174, 762);
            this.btnRemove.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.Size = new System.Drawing.Size(107, 38);
            this.btnRemove.TabIndex = 4;
            this.btnRemove.Text = "Remove";
            this.btnRemove.UseVisualStyleBackColor = true;
            this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(406, 762);
            this.btnSave.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(107, 38);
            this.btnSave.TabIndex = 5;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnCreateGarantee
            // 
            this.btnCreateGarantee.Location = new System.Drawing.Point(14, 762);
            this.btnCreateGarantee.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnCreateGarantee.Name = "btnCreateGarantee";
            this.btnCreateGarantee.Size = new System.Drawing.Size(143, 38);
            this.btnCreateGarantee.TabIndex = 6;
            this.btnCreateGarantee.Text = "Tạo Bảo Lãnh";
            this.btnCreateGarantee.UseVisualStyleBackColor = true;
            this.btnCreateGarantee.Click += new System.EventHandler(this.btnCreateGarantee_Click);
            // 
            // txbGaranteeValue
            // 
            this.txbGaranteeValue.Location = new System.Drawing.Point(200, 508);
            this.txbGaranteeValue.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txbGaranteeValue.Name = "txbGaranteeValue";
            this.txbGaranteeValue.Size = new System.Drawing.Size(141, 31);
            this.txbGaranteeValue.TabIndex = 7;
            this.txbGaranteeValue.Text = "0";
            this.txbGaranteeValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(17, 517);
            this.label12.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(136, 25);
            this.label12.TabIndex = 3;
            this.label12.Text = "Giá Trị Bảo Lãnh";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(14, 567);
            this.label13.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(192, 25);
            this.label13.TabIndex = 3;
            this.label13.Text = "Ngày hết hạn bảo lãnh";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label14.Location = new System.Drawing.Point(343, 513);
            this.label14.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(31, 30);
            this.label14.TabIndex = 3;
            this.label14.Text = "%";
            // 
            // dateTimePickerDateSignedPO
            // 
            this.dateTimePickerDateSignedPO.CustomFormat = "dd/mm/yyyy";
            this.dateTimePickerDateSignedPO.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePickerDateSignedPO.Location = new System.Drawing.Point(350, 208);
            this.dateTimePickerDateSignedPO.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dateTimePickerDateSignedPO.Name = "dateTimePickerDateSignedPO";
            this.dateTimePickerDateSignedPO.Size = new System.Drawing.Size(284, 31);
            this.dateTimePickerDateSignedPO.TabIndex = 8;
            // 
            // dateTimePickerDurationDateContract
            // 
            this.dateTimePickerDurationDateContract.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePickerDurationDateContract.Location = new System.Drawing.Point(350, 258);
            this.dateTimePickerDurationDateContract.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dateTimePickerDurationDateContract.Name = "dateTimePickerDurationDateContract";
            this.dateTimePickerDurationDateContract.Size = new System.Drawing.Size(284, 31);
            this.dateTimePickerDurationDateContract.TabIndex = 9;
            // 
            // dateTimePickerActiveDateContract
            // 
            this.dateTimePickerActiveDateContract.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePickerActiveDateContract.Location = new System.Drawing.Point(350, 358);
            this.dateTimePickerActiveDateContract.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dateTimePickerActiveDateContract.Name = "dateTimePickerActiveDateContract";
            this.dateTimePickerActiveDateContract.Size = new System.Drawing.Size(284, 31);
            this.dateTimePickerActiveDateContract.TabIndex = 10;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(14, 17);
            this.label16.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(167, 25);
            this.label16.TabIndex = 11;
            this.label16.Text = "Kế Hoạch Mua Sắm";
            // 
            // txbKHMS
            // 
            this.txbKHMS.Location = new System.Drawing.Point(200, 8);
            this.txbKHMS.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txbKHMS.Name = "txbKHMS";
            this.txbKHMS.Size = new System.Drawing.Size(434, 31);
            this.txbKHMS.TabIndex = 12;
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(570, 417);
            this.label17.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(49, 25);
            this.label17.TabIndex = 13;
            this.label17.Text = "VND";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(570, 467);
            this.label18.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(54, 25);
            this.label18.TabIndex = 14;
            this.label18.Text = "Ngày";
            // 
            // ExpirationDate
            // 
            this.ExpirationDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.ExpirationDate.Location = new System.Drawing.Point(350, 560);
            this.ExpirationDate.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.ExpirationDate.Name = "ExpirationDate";
            this.ExpirationDate.Size = new System.Drawing.Size(284, 31);
            this.ExpirationDate.TabIndex = 15;
            // 
            // txbGaranteeActiveDate
            // 
            this.txbGaranteeActiveDate.Location = new System.Drawing.Point(200, 560);
            this.txbGaranteeActiveDate.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txbGaranteeActiveDate.Name = "txbGaranteeActiveDate";
            this.txbGaranteeActiveDate.Size = new System.Drawing.Size(111, 31);
            this.txbGaranteeActiveDate.TabIndex = 16;
            this.txbGaranteeActiveDate.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // ContractInfoChildForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(651, 917);
            this.Controls.Add(this.txbGaranteeActiveDate);
            this.Controls.Add(this.ExpirationDate);
            this.Controls.Add(this.label18);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.txbKHMS);
            this.Controls.Add(this.label16);
            this.Controls.Add(this.dateTimePickerActiveDateContract);
            this.Controls.Add(this.dateTimePickerDurationDateContract);
            this.Controls.Add(this.dateTimePickerDateSignedPO);
            this.Controls.Add(this.txbGaranteeValue);
            this.Controls.Add(this.btnCreateGarantee);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnRemove);
            this.Controls.Add(this.btnEdit);
            this.Controls.Add(this.btnNewPO);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.tbxSiteB);
            this.Controls.Add(this.tbxSiteA);
            this.Controls.Add(this.tbxDurationPO);
            this.Controls.Add(this.tbxValueContract);
            this.Controls.Add(this.txbTypeContract);
            this.Controls.Add(this.tbxDurationContract);
            this.Controls.Add(this.tbAccountingCode);
            this.Controls.Add(this.tbBidName);
            this.Controls.Add(this.tbContract);
            this.Controls.Add(this.btnDescriptionA);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "ContractInfoChildForm";
            this.Text = "ContractInfoChildForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnDescriptionA;
        private System.Windows.Forms.TextBox tbContract;
        private System.Windows.Forms.TextBox tbBidName;
        private System.Windows.Forms.TextBox tbAccountingCode;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox tbxDurationContract;
        private System.Windows.Forms.TextBox txbTypeContract;
        private System.Windows.Forms.TextBox tbxValueContract;
        private System.Windows.Forms.TextBox tbxDurationPO;
        private System.Windows.Forms.TextBox tbxSiteA;
        private System.Windows.Forms.TextBox textBox10;
        private System.Windows.Forms.Button btnNewPO;
        private System.Windows.Forms.Button btnEdit;
        private System.Windows.Forms.Button btnRemove;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.TextBox tbxSiteB;
        private System.Windows.Forms.Button btnCreateGarantee;
        private System.Windows.Forms.TextBox txbGaranteeValue;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.DateTimePicker dateTimePickerDateSignedPO;
        private System.Windows.Forms.DateTimePicker dateTimePickerDurationDateContract;
        private System.Windows.Forms.DateTimePicker dateTimePickerActiveDateContract;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.TextBox txbKHMS;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.DateTimePicker ExpirationDate;
        private System.Windows.Forms.TextBox txbGaranteeActiveDate;
    }
}