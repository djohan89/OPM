﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace OPM.GUI
{
    public partial class DeliverPartInforDetail : Form
    {
        public delegate void UpdateCatalogDelegate(string value);
        public UpdateCatalogDelegate UpdateCatalogPanel;
        public DeliverPartInforDetail()
        {
            InitializeComponent();
        }
    }
}
