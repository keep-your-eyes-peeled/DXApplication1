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
using ClassLibrary1;
using DevExpress.XtraSplashScreen;
using EasyDox;

namespace DXApplication1
{
    public partial class Login : Form
    {
        public Login()
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            InitializeComponent();
            SplashScreenManager.CloseForm();
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
                string post = this.dataSet11.USERS.FillBy(this.dataSet11.MyOraConnection, hesh);
                string emplnum = this.dataSet11.USERS.FindENUM(myOraConnection, hesh);
                switch (this.dataSet11.USERS.FindByHash(myOraConnection, hesh))
                {
                    case true:
                        Glavn glavn = new Glavn();
                        glavn.post = post;
                        glavn.emplnum = emplnum;
                        glavn.Show();
                        this.Hide();
                        break;
                    case false:
                        MessageBox.Show("Ошибка авторизации!");
                        break;
                }
            }
        }



        private void Login_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();

        }
    }
}
