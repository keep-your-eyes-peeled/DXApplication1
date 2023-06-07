using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;
using System.Diagnostics;

namespace DXApplication1
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            if (textEdit1.Text.Length == 0 || textEdit2.Text.Length == 0)
            {
                MessageBox.Show("Введите данные для авторизации!");
            }
            else
            {
                this.dataSet11.SetOracleConnectionString();
                string hesh = this.dataSet11.GetStringHashSum(textEdit1.Text, textEdit2.Text);
                OracleConnection myOraConnection= this.dataSet11.MyOraConnection;                
                switch (this.dataSet11.USERS.FindByHash(myOraConnection, hesh))
                {
                    case true:
                        Glavn glavn = new Glavn();
                        glavn.Show();
                        this.Hide();
                        break;
                    case false:
                        MessageBox.Show("Ошибка авторизации!");
                        break;
                }
                
                /*if (this.dataSet11.USERS.CheckHashSum(this.dataSet11.OraConnection,
                    this.dataSet11.GetStringHashSum(textEdit1.Text, textEdit2.Text)) == "Программист")
                {
                    Form1 form1 = new Form1();
                    form1.Show();
                    this.Hide();
                }
                else if (this.dataSet11.USERS.CheckHashSum(this.dataSet11.OraConnection,
                    this.dataSet11.GetStringHashSum(textEdit1.Text, textEdit2.Text)) == "Бухгалтер")
                {
                    Form2 form2 = new Form2();
                    form2.Show();
                    this.Hide();
                }*/
                /*else
                {
                    MessageBox.Show("Ошибка авторизации!");
                }*/
            }
        }



        private void Login_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
            //this.Close();
        }
    }
}
