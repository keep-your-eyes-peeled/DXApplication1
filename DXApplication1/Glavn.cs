using BitMiracle.LibTiff.Classic;
using DevExpress.XtraBars.Docking2010;
using DevExpress.XtraBars.Docking2010.Views.NativeMdi;
using EasyDox;
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
using GroupDocs.Conversion.Options.Convert;
using GroupDocs.Conversion;
using System.IO;
using Microsoft.Office.Interop.Word;
using DevExpress.XtraPrinting.Preview.Native;
using DevExpress.XtraSplashScreen;
using DevExpress.XtraReports.UI;
using System.Data.OracleClient;

namespace DXApplication1
{
    public partial class Glavn : Form
    {
        public Form1 form1 = null;
        public string post = null;
        public string emplnum = null;
        Form2 form2 = null;
        Form4 form4 = null;
        ClassLibrary1.Report report=new ClassLibrary1.Report(new Dictionary<string, string>());
        private int counter = 1;
        public int Counter { get => counter; set => counter = value; }
        public Glavn()
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));            
            InitializeComponent();
            SplashScreenManager.CloseForm();
            
        }

        private void Glavn_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void barButtonItem3_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            report = form1.Report;
            var engine = new Engine();
            engine.Merge("Report_temp.docx", report.FieldValues, "Report_out.docx");
            SplashScreenManager.CloseForm();
            
            //Process.Start("Report_out.docx");
            /*engine.Merge(@"D:\work\Диплом\Документы\Report_temp.docx", report.FieldValues, @"D:\work\Диплом\Документы\Report_out.docx");
            Process.Start(@"D:\work\Диплом\Документы\Report_out.docx");*/
        }

        private void barButtonItem4_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            report = form1.Report;
            var engine = new Engine();
            engine.Merge("Report_temp.docx", report.FieldValues, "Report_out.docx");
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            FileInfo file = new FileInfo("Report_out.pdf");
            word.Visible = false;
            word.ScreenUpdating = false;
            object oMissing = System.Reflection.Missing.Value;
            object filename = (object)file.FullName;
            Microsoft.Office.Interop.Word.Document doc = word.Documents.Open(ref filename, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing);
            doc.SaveAs2("Report_out.pdf", WdSaveFormat.wdFormatPDF,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                 ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            ((Microsoft.Office.Interop.Word._Document)doc).Close(WdSaveOptions.wdSaveChanges, ref oMissing, ref oMissing);
            ((Microsoft.Office.Interop.Word._Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
            word = null;
            SplashScreenManager.CloseForm();
            
        }

        private void barButtonItem5_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            report = form1.Report;
            report.CreateExcelFile();
            SplashScreenManager.CloseForm();
        }

        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            form1 = new Form1();
            form1.Text = "Отчет " + Counter++;
            form1.MdiParent = this;
            form1.emplnum = emplnum;
            form1.Show();
            SplashScreenManager.CloseForm();            
        }

        private void Glavn_FormClosing(object sender, FormClosingEventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }



        private void barButtonItem2_ItemClick_1(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            /*SplashScreenManager.ShowForm(typeof(WaitForm1));
            form1 = new Form1();
            report.OpenExcelFile();
            form1.Report = report;            
            form1.MdiParent = this;
            form1.Show();
            
            SplashScreenManager.CloseForm();*/
        }

        private void barButtonItem6_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            if (form4 == null)
            {
                form4 = new Form4(post);
                form4.FormClosed += instanceHasBeenClosed;
                form4.emplnum = emplnum;
                form4.MdiParent = this;
                form4.Show();
            }
            else
            {
                form4.BringToFront();
            }
            SplashScreenManager.CloseForm();            
        }
        

        private void barButtonItem7_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            if (form2 == null)
            {
                form2 = new Form2(post);
                form2.FormClosed += instanceHasBeenClosed;
                form2.MdiParent = this;
                form2.emplnum = emplnum;
                form2.Show();
            }
            else
            {
                form2.BringToFront();
            }
            SplashScreenManager.CloseForm();
        }

        private void Glavn_Load(object sender, EventArgs e)
        {
            /*SplashScreenManager.ShowForm(typeof(WaitForm1));
            form1 = new Form1();
            form1.Text = "Отчет " + Counter++;
            form1.MdiParent = this;
            form1.Show();
            SplashScreenManager.CloseForm();*/
        }

        private void instanceHasBeenClosed(object sender, FormClosedEventArgs e)
        {
            switch (sender.GetType().Name)
            {
                case "Form4":
                    form4 = null;
                    break;
                case "Form2":
                    form2 = null;
                    break;
                default:
                    return;
            }
        }
    }
}
