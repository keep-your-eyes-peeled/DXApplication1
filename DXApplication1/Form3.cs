using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using ClassLibrary1;
using DevExpress.ChartRangeControlClient.Core;
using System.Security.Cryptography;

namespace DXApplication1
{
    public partial class Form3 : Form
    {
        public DataSet1.INCDOCDataTable IncDoc = new DataSet1.INCDOCDataTable();
        public DataSet1.INCDOCSDataTable docs = null;
        string repnum = "";
        public string emplnum = null;
        string fileName = null;
        string filePath = null;
        public Form3(string _repnum, DataSet1.INCDOCSDataTable _docs)
        {
            InitializeComponent();
            this.repnum = _repnum;
            this.docs = _docs;
        }


        private void simpleButton1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if(openFileDialog.ShowDialog() == DialogResult.OK )
            {
                fileName = System.IO.Path.GetFileName(openFileDialog.FileName);
                filePath = System.IO.Path.GetFullPath(openFileDialog.FileName);
                labelControl1.Text = openFileDialog.FileName;                                          
            }
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.BindingSource bs2 = new System.Windows.Forms.BindingSource();
            bs2.DataSource = docs;
            bs2.AddNew();

            DataSet1.INCDOCSRow docsRow = ((DataRowView)bs2.Current).Row
                as DataSet1.INCDOCSRow;

            docsRow.PATH = labelControl1.Text;
            docsRow.REPNUM = this.repnum;
            docsRow.CREATORENUM = emplnum;


            docs.AddINCDOCSRow(((DataSet1.INCDOCSRow)((DataRowView)bs2.Current).Row));
            bs2.EndEdit();

            bindingSource1.DataSource = this.IncDoc;
            bindingSource1.AddNew();

            DataSet1.INCDOCRow curRow = ((DataRowView)bindingSource1.Current).Row
                as DataSet1.INCDOCRow;            

            curRow.Date = textEdit1.Text;
            curRow.Number = textEdit2.Text;
            curRow.Title = textEdit3.Text;
            curRow.Sum1 = textEdit5.Text;
            curRow.Sum2 = textEdit6.Text;
            curRow.Debet = textEdit4.Text;
            curRow.Kredit = textEdit7.Text;

            IncDoc.AddINCDOCRow(((DataSet1.INCDOCRow)((DataRowView)bindingSource1.Current).Row));
            bindingSource1.EndEdit();

            if (File.Exists($@"D:\Документы\Входящие документы\{fileName}") == false)
            {
                File.Copy(filePath, $@"D:\Документы\Входящие документы\{fileName}");
            }
            else
            {
                File.Copy(filePath, $@"D:\Документы\Входящие документы\(copy{new Random().Next(1,100000)}){fileName}");
            }


            this.Close();
        }
    }
}
