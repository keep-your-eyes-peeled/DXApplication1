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

namespace DXApplication1
{
    public partial class Form3 : Form
    {
        private List<object> item;
        public List<object> Item { get => item; set => item = value; }
        public Form3()
        {
            InitializeComponent();
        }


        private void simpleButton1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if(openFileDialog.ShowDialog() == DialogResult.OK )
            {
                labelControl1.Text = openFileDialog.FileName;
            }
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            Item = new List<object> { textEdit1.Text, textEdit2.Text, textEdit3.Text, 
                textEdit5.Text, textEdit6.Text, textEdit4.Text};
            this.Close();
        }
    }
}
