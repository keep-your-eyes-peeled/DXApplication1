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
    public partial class Form4 : Form
    {
        public string emplnum = null;
        string post = null;
        public Form4(string post)
        {
            InitializeComponent();
            this.post = post;
        }

        private void Form4_Load(object sender, EventArgs e)
        {
            if (post != "Главный бухгалтер" && post != "Бухгалтер")
            {
                this.dataSet11.INCDOCS.FillBy(this.dataSet11.MyOraConnection, emplnum);
            }
            else
            {
                this.dataSet11.INCDOCS.Fill(this.dataSet11.MyOraConnection);
            }
        }

        private void gridView1_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            if (e.Column.Name == "colPATH")
            {
                Process.Start(e.CellValue.ToString());
            }
        }
    }
}
