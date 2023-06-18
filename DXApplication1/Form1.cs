using System;
using System.Collections.Generic;
using ClassLibrary1;
using System.Linq;
using System.Text.RegularExpressions;
using System.Data;
using DevExpress.ClipboardSource.SpreadsheetML;
using DevExpress.XtraSplashScreen;
using System.Windows.Forms;
using DevExpress.ChartRangeControlClient.Core;
using DevExpress.XtraSpellChecker.Parser;
using System.Diagnostics.Metrics;
using DevExpress.XtraEditors;
using System.IO;
using EasyDox;
using System.Drawing;

namespace DXApplication1
{    
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        
        private ClassLibrary1.Report report;
        
        public Report Report { get => report; set => report = value; }

        public string emplnum =null;
        public string post = null;
        

        


        enum mesyac { Января, Февраля, Марта, Апреля, Мая, Июня, Июля, Августа, Сентября, Октября, Ноября, Декабря };
        public Form1()
        {
            InitializeComponent();
        }

        public ClassLibrary1.Report GetReport()
        {            
            //Console.WriteLine(Сумма.Пропись((double)2340, Валюта.Рубли));

            DataSet1.STAFFDataTable staffTable = new DataSet1.STAFFDataTable();
            staffTable.FindSotr(this.dataSet11.MyOraConnection, textBox2.Text);
            System.Data.DataRow row = staffTable.Rows[0];
            //MessageBox.Show(((DataSet1.STAFFRow)row).NAME);
            //DataSet1.STAFFRow row = staffTable.Rows[0];
            /*foreach (DataSet1.STAFFRow row in staffTable.Rows[0])
            {
                MessageBox.Show(row.NAME);
            }*/

            string name = Regex.Replace(((DataSet1.STAFFRow)row).NAME, @"[^\w ]", "");
            name = Regex.Replace(name, @"(\w+)\s(\w+)\s(\w+)", m => string.Format("{0} {1}. {2}.",
                m.Groups[1], m.Groups[2].Value.FirstOrDefault(), m.Groups[3].Value.FirstOrDefault()), RegexOptions.IgnoreCase);


            //MessageBox.Show(((System.Data.DataRowView)bindingSource1[0])[4].ToString());

            float sum = 0;
            foreach (System.Data.DataRowView dataRow in bindingSource1)
            {
                sum += float.Parse(dataRow[4].ToString());
            }
            //MessageBox.Show(sum.ToString());

            string sumProp = Сумма.Пропись(sum, Валюта.Рубли);
            string sumPropRub = "";
            string sumPropKop = "";
            int index = 0;
            foreach (string str in sumProp.Split(' '))
            {
                if (str.StartsWith("руб"))
                {
                    break;
                }
                index++;
            }
            for (int i = 0; i < index; i++)
            {
                if ((index - i) != 1)
                {
                    sumPropRub += sumProp.Split(' ')[i] + " ";
                }
                else
                {
                    sumPropRub += sumProp.Split(' ')[i];
                }
            }
            sumPropKop += sumProp.Split(' ')[index + 1];

            //MessageBox.Show(sumPropRub);
            //MessageBox.Show(sumPropKop);

            string sumPropT = "";
            string sumPropOst = "";
            int index1 = 0;
            foreach (string str in sumProp.Split(' '))
            {
                if (str.StartsWith("тыс"))
                {
                    break;
                }
                index1++;
            }
            for (int i = 0; i < index1; i++)
            {
                if ((index1 - i) != 1)
                {
                    sumPropT += sumProp.Split(' ')[i] + " ";
                }
                else
                {
                    sumPropT += sumProp.Split(' ')[i];
                }
            }
            if (index1 == sumProp.Split(' ').Length)
            {
                sumPropT = "";
                index1 = -1;
            }
            else
            {
                sumPropT += " " + sumProp.Split(' ')[index1];
            }

            for (int i = index1 + 1; i < index; i++)
            {
                if ((index1 - i) != 1)
                {
                    sumPropOst += sumProp.Split(' ')[i] + " ";
                }
                else
                {
                    sumPropOst += sumProp.Split(' ')[i];
                }
                //MessageBox.Show(sumPropOst);
            }
            //MessageBox.Show(sumPropOst);
            //MessageBox.Show(sumPropT + sumProp.Split(' ')[index1]);
            //MessageBox.Show(sumPropOst);

            var staffTable1 = new DataSet1.STAFFDataTable();
            staffTable1.FindDirector(this.dataSet11.MyOraConnection);
            var row1 = staffTable1.Rows[0];

            string name1 = Regex.Replace(((DataSet1.STAFFRow)row1).NAME, @"[^\w ]", "");
            name1 = Regex.Replace(name1, @"(\w+)\s(\w+)\s(\w+)", m => string.Format("{0} {1}. {2}.",
                m.Groups[1], m.Groups[2].Value.FirstOrDefault(), m.Groups[3].Value.FirstOrDefault()), RegexOptions.IgnoreCase);

            var staffTable2 = new DataSet1.STAFFDataTable();
            staffTable2.FindGlavBuh(this.dataSet11.MyOraConnection);
            var row2 = staffTable2.Rows[0];

            string name2 = Regex.Replace(((DataSet1.STAFFRow)row2).NAME, @"[^\w ]", "");
            name2 = Regex.Replace(name2, @"(\w+)\s(\w+)\s(\w+)", m => string.Format("{0} {1}. {2}.",
                m.Groups[1], m.Groups[2].Value.FirstOrDefault(), m.Groups[3].Value.FirstOrDefault()), RegexOptions.IgnoreCase);

            //DateTime dateTime = DateTime.TryParse(textEdit3.DateTime, dateTime)
            //MessageBox.Show(textEdit3.DateTime.Month.ToString());
            string pererashod = "0";
            if (sum > float.Parse(textEdit1.Text))
            {
                pererashod = (sum - float.Parse(textEdit1.Text)).ToString();
            }
            string year = textEdit3.DateTime.Year.ToString().Substring(2, 2);
            string year1 = textEdit2.DateTime.Year.ToString().Substring(2, 2);

            List<DataRowView> rows = new List<DataRowView> { };
            foreach (System.Data.DataRowView dataRow in bindingSource1)
            {
                rows.Add(dataRow);
            }

            string sum1 = "";
            //string sum2 = "";
            if (sum > float.Parse(textEdit1.Text))
            {
                sum1 = pererashod;
            }
            else if (sum < float.Parse(textEdit1.Text))
            {
                sum1 = (float.Parse(textEdit1.Text) - sum).ToString();
            }
            else
            {
                sum1 = "0";
            }
            
            Dictionary<string, string> dict = new Dictionary<string, string> 
            {
                {"Номер", textBox1.Text},
                {"Дата", textEdit3.DateTime.Date.ToString("dd.MM.yyyy")},
                {"Код", "-"},
                {"ТабНом", textBox2.Text},
                {"СтруктурноеПодразделение", ((DataSet1.STAFFRow)row).STRDIV},
                {"ФамилияИнициалы", name},
                {"проф", ((DataSet1.STAFFRow)row).POST},
                {"НазначениеАванса", textBox111.Text},
                {"Остаток", textEdit4.Text},
                {"Перерасход", textEdit5.Text},
                {"ИтогПолуч", textEdit1.Text},
                {"Израсход", sum.ToString()},
                {"Остаток1", (float.Parse(textEdit1.Text)-sum).ToString()},
                {"Перерасход1", pererashod},
                {"Прил", bindingSource1.Count.ToString()},
                {"листах", bindingSource1.Count.ToString()},
                {"сумпропис", sumPropRub},
                {"Руб", sum.ToString()},
                {"Коп", sumPropKop},
                //{"НаимОрг", "Рога и Копыта"},
                {"Тысячи", sumPropT},
                {"СотДесЕд", sumPropOst},
                //{"Долж", "Главный бухгалтер"},
                {"РасшПодп", name1},
                {"дд", textEdit3.DateTime.Day.ToString()},
                {"мм1", ((mesyac)(textEdit3.DateTime.Month - 1)).ToString()},
                {"г2", year},
                {"nu", textBox1.Text},
                {"Н", ""},
                {"ОТ", ""},
                {"МЕС", ""},
                {"г5", year1},
                {"ДЕ1", ""},
                {"БУХ", ""},
                {"г1", year1},
                {"Де", textEdit2.DateTime.Day.ToString()},
                {"Месяц", ((mesyac)(textEdit2.DateTime.Month - 1)).ToString()},
                {"го", year1},
                {"ГлБухРасшПодп", name2},
                {"БухРасшПодп", name2},
                {"ПолАва", textEdit1.Text},
                {"ПолВал", ""},
                {"СУ", ""},
                {"СУС", ""},

                /*{"Сч1", "11"},
                {"Сч2", "12"},
                {"Сч3", "13"},
                {"Сч4", "14"},
                {"Сч5", "15"},
                {"Сч6", "16"},
                {"Сч7", "17"},
                {"Сч8", "18"},
                {"Сч9", "19"},
                {"Сч10", "20"},
                {"Сч11", "21"},
                {"Сч12", "22"},
                {"Сч13", "23"},
                {"Сч14", "24"},
                {"Сч15", "25"},
                {"Сч16", "26"},
                {"Сум1", "11"},
                {"Сум2", "12"},
                {"Сум3", "13"},
                {"Сум4", "14"},
                {"Сум5", "15"},
                {"Сум6", "16"},
                {"Сум7", "17"},
                {"Сум8", "18"},
                {"Сум9", "11"},
                {"Сум10", "12"},
                {"Сум11", "13"},
                {"Сум12", "14"},
                {"Сум13", "15"},
                {"Сум14", "16"},
                {"Сум15", "17"},
                {"Сум16", "18"},
                {"пов9", "17"},
                {"пов10", "18"},
                {"рп1", "Крюков"},
                {"дс8", "18"},
                {"н", "22"},
                {"от1", "18"},
                {"мес1", "май"},
                {"г1", "23"},*/

                {"ит1", sum.ToString()},
                {"ит2", "0"},
                {"ит3", sum.ToString()},
                {"ит4", "0"}
            };

            int counter = 1;
            int _counter = 1;
            Dictionary<string, string> sch = new Dictionary<string, string> { };
            foreach (DataRowView DRrow in rows)
            {
                float _sum = 0;
                if (!sch.ContainsValue(DRrow[5].ToString()))
                {
                    sch.Add("Сч" + _counter, DRrow[5].ToString());
                    sch.Add("Сч" + (_counter + 8), DRrow[5].ToString() + "0");
                    foreach (System.Data.DataRowView dataRow in bindingSource1)
                    {
                        if (dataRow.Row.ItemArray[5].ToString() == sch["Сч" + _counter])
                        {
                            _sum += float.Parse(dataRow.Row.ItemArray[3].ToString());
                        }
                    }
                    sch.Add("Сум" + _counter, _sum.ToString());
                    sch.Add("Сум" + (_counter + 8), _sum.ToString());
                    _counter++;
                }
                _sum = 0;
                dict.Add("пп" + counter, counter.ToString());
                dict.Add("д" + counter, DRrow[0].ToString());
                dict.Add("нн" + counter, DRrow[1].ToString());
                dict.Add("нд" + counter, DRrow[2].ToString());
                dict.Add("по" + counter, DRrow[3].ToString());
                dict.Add("пов" + counter, "0");
                dict.Add("пк" + counter, DRrow[4].ToString());
                dict.Add("пкв" + counter, "0");
                dict.Add("дс" + counter, DRrow[5].ToString());
                counter++;
            }


            for (int i = (sch.Count / 4) + 1; i < 9; i++)
            {
                dict.Add("Сч" + i, "");
                dict.Add("Сум" + i, "");
                dict.Add("Сч" + (i + 8), "");
                dict.Add("Сум" + (i + 8), "");
                //counter++;
            }

            foreach (var kvp in sch)
            {
                if (!dict.ContainsKey(kvp.Key))
                {
                    dict.Add(kvp.Key, kvp.Value);
                }
            }
            for (int i = rows.Count + 1; i < 11; i++)
            {
                dict.Add("пп" + counter, "");
                dict.Add("д" + counter, "");
                dict.Add("нн" + counter, "");
                dict.Add("нд" + counter, "");
                dict.Add("по" + counter, "");
                dict.Add("пов" + counter, "");
                dict.Add("пк" + counter, "");
                dict.Add("пкв" + counter, "");
                dict.Add("дс" + counter, "");
                counter++;
            }


            Report = new ClassLibrary1.Report(dict);
            return Report;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (post == "Главный бухгалтер")
            {
                textEdit2.Enabled = true;
                layoutControlItem9.Enabled = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            

        }

        private void Form1_FormClosed(object sender, System.Windows.Forms.FormClosedEventArgs e)
        {
            //System.Windows.Forms.Application.Exit();
            //this.Close();
            //Process.GetCurrentProcess().Kill();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            
            Form3 form3 = new Form3(textBox1.Text, this.dataSet11.INCDOCS);
            form3.emplnum = emplnum;
            form3.ShowDialog();
            
            this.dataSet11.INCDOC.Merge(form3.IncDoc);
            SplashScreenManager.CloseForm();
            
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            if (bindingSource1.Count == 0)
            {
                MessageBox.Show("Вы не внесли входящие документы!");
                return;
            }
            GetReport();
            var engine = new Engine();
            engine.Merge("Report_temp.docx", report.FieldValues, "Report_out.docx");
            System.Windows.Forms.BindingSource bs1 = new System.Windows.Forms.BindingSource();
            bs1.DataSource = this.dataSet11.REPORTS;
            bs1.AddNew();
            if (File.Exists($@"D:\Документы\Отчеты\{report.GetReportName()}") == false)
            {
                File.Copy("Report_out.docx", $@"D:\Документы\Отчеты\{report.GetReportName()}");
                File.Delete("Report_out.docx");
            }
            else
            {
                File.Copy("Report_out.docx", $@"D:\Документы\Отчеты\(copy{new Random().Next(1, 100000)}){report.GetReportName()}");
                File.Delete("Report_out.docx");
            }
            DataSet1.REPORTSRow Row = ((DataRowView)bs1.Current).Row
            as DataSet1.REPORTSRow;
            Row.NUM = report.FieldValues["Номер"];
            Row.PATH = $@"D:\Документы\Отчеты\{report.GetReportName()}";
            Row.CREATORENUM = emplnum;
            Row.APPROVED = false;
            this.dataSet11.REPORTS.AddREPORTSRow(((DataSet1.REPORTSRow)((DataRowView)bs1.Current).Row));
            bs1.EndEdit();
            this.dataSet11.INCDOCS.Update(this.dataSet11.MyOraConnection);
            this.dataSet11.REPORTS.Update(this.dataSet11.MyOraConnection);
            this.Close();
            SplashScreenManager.CloseForm();
        }
    }
    
}
