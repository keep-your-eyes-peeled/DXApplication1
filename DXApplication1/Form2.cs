﻿using ClassLibrary1;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DXApplication1
{
    public partial class Form2 : Form
    {
        public string emplnum = null;
        string post = null;
        public Form2(string post)
        {
            InitializeComponent();
            this.post = post;
        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            //Application.Exit();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            /*OpenFileDialog OFD = new OpenFileDialog();
            if (OFD.ShowDialog() == DialogResult.OK)
            {
                MessageBox.Show(OFD.FileName);
            }*/
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            if(post!="Главный бухгалтер" && post != "Бухгалтер")
            {
                this.dataSet11.REPORTS.FillBy(this.dataSet11.MyOraConnection, emplnum);                
            }
            else
            {
                this.dataSet11.REPORTS.Fill(this.dataSet11.MyOraConnection);
            }
            if(post!="Главный бухгалтер")
            {
                dataSet11.REPORTS.APPROVEDColumn.ReadOnly = true;
            }
        }

        private void gridView1_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            if (e.Column.Name == "colPATH")
            {
                Process.Start(e.CellValue.ToString());
            }
        }

        private void simpleButton1_Click_1(object sender, EventArgs e)
        {
            this.dataSet11.REPORTS.Update(this.dataSet11.MyOraConnection);
        }

        private void gridView1_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            
        }
    }
}
