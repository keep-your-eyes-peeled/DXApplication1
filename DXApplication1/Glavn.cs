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

namespace DXApplication1
{
    public partial class Glavn : Form
    {
        public Form1 form1;
        ClassLibrary1.Report report=new ClassLibrary1.Report(new Dictionary<string, string>());
        public Glavn()
        {
            InitializeComponent();            
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
            engine.Merge(@"D:\work\Диплом\Документы\Report_temp.docx", report.FieldValues, @"D:\work\Диплом\Документы\Report_out.docx");
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            FileInfo file = new FileInfo(@"D:\work\Диплом\Документы\Report_out.docx");
            word.Visible = false;
            word.ScreenUpdating = false;
            object oMissing = System.Reflection.Missing.Value;
            object filename = (object)file.FullName;
            Microsoft.Office.Interop.Word.Document doc = word.Documents.Open(ref filename, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing);
            doc.SaveAs2(@"D:\work\Диплом\Документы\Report_out.pdf", WdSaveFormat.wdFormatPDF,
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
            form1.MdiParent = this;
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
    }
}
