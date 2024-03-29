﻿using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Microsoft.CSharp.RuntimeBinder;

namespace ClassLibrary1
{
    public class Report
    {
        public Report(Dictionary<string, string> fieldValues)
        {
            FieldValues = fieldValues;
        }

        #region Fields
        private Dictionary<string, string> fieldValues;
        private Excel.Application ex;
        private Excel.Workbook workbook;
        private Excel.Worksheet worksheet;
        #endregion

        #region Properties
        public Dictionary<string, string> FieldValues { get => fieldValues; set => fieldValues = value; }
        public Excel.Application Ex { get => ex; set => ex = value; }
        public Workbook Workbook { get => workbook; set => workbook = value; }
        public Worksheet Worksheet { get => worksheet; set => worksheet = value; }
        #endregion

        #region Methods
        public void CreateExcelFile()
        {
            try
            {
                Ex = new Microsoft.Office.Interop.Excel.Application();
                Workbook = Ex.Workbooks.Open(Environment.CurrentDirectory + "Report_out.xlsx");
                Worksheet = Workbook.Worksheets["Лист1"];

                Worksheet.Range["BB21"].Value = fieldValues["ТабНом"];
                Worksheet.Range["AV11"].Value = fieldValues["Тысячи"];
                Worksheet.Range["AN12"].Value = fieldValues["СотДесЕд"];
                Worksheet.Range["X14"].Value = fieldValues["Номер"];
                Worksheet.Range["AE14"].Value = fieldValues["Дата"];
                Worksheet.Range["AX15"].Value = fieldValues["РасшПодп"];
                Worksheet.Range["AO17"].Value = fieldValues["Де"];
                Worksheet.Range["AS17"].Value = fieldValues["Месяц"];
                Worksheet.Range["BE17"].Value = fieldValues["го"];
                Worksheet.Range["P20"].Value = fieldValues["СтруктурноеПодразделение"];
                Worksheet.Range["K21"].Value = fieldValues["ФамилияИнициалы"];
                Worksheet.Range["BB20"].Value = fieldValues["Код"];
                Worksheet.Range["M24"].Value = fieldValues["проф"];
                Worksheet.Range["AQ24"].Value = fieldValues["НазначениеАванса"];
                Worksheet.Range["X27"].Value = fieldValues["Остаток"];
                Worksheet.Range["X28"].Value = fieldValues["Перерасход"];
                Worksheet.Range["X29"].Value = fieldValues["ПолАва"];
                Worksheet.Range["X30"].Value = fieldValues["ПолВал"];
                Worksheet.Range["X33"].Value = fieldValues["ИтогПолуч"];
                Worksheet.Range["X34"].Value = fieldValues["Израсход"];
                Worksheet.Range["X35"].Value = fieldValues["Остаток1"];
                Worksheet.Range["X36"].Value = fieldValues["Перерасход1"];
                Worksheet.Range["AG29"].Value = fieldValues["Сч1"];
                Worksheet.Range["AG30"].Value = fieldValues["Сч2"];
                Worksheet.Range["AG31"].Value = fieldValues["Сч3"];
                Worksheet.Range["AG32"].Value = fieldValues["Сч4"];
                Worksheet.Range["AG33"].Value = fieldValues["Сч5"];
                Worksheet.Range["AG34"].Value = fieldValues["Сч6"];
                Worksheet.Range["AG35"].Value = fieldValues["Сч7"];
                Worksheet.Range["AG36"].Value = fieldValues["Сч8"];
                Worksheet.Range["AO29"].Value = fieldValues["Сум1"];
                Worksheet.Range["AO30"].Value = fieldValues["Сум2"];
                Worksheet.Range["AO31"].Value = fieldValues["Сум3"];
                Worksheet.Range["AO32"].Value = fieldValues["Сум4"];
                Worksheet.Range["AO33"].Value = fieldValues["Сум5"];
                Worksheet.Range["AO34"].Value = fieldValues["Сум6"];
                Worksheet.Range["AO35"].Value = fieldValues["Сум7"];
                Worksheet.Range["AO36"].Value = fieldValues["Сум8"];
                Worksheet.Range["AW29"].Value = fieldValues["Сч9"];
                Worksheet.Range["AW30"].Value = fieldValues["Сч10"];
                Worksheet.Range["AW31"].Value = fieldValues["Сч11"];
                Worksheet.Range["AW32"].Value = fieldValues["Сч12"];
                Worksheet.Range["AW33"].Value = fieldValues["Сч13"];
                Worksheet.Range["AW34"].Value = fieldValues["Сч14"];
                Worksheet.Range["AW35"].Value = fieldValues["Сч15"];
                Worksheet.Range["AW36"].Value = fieldValues["Сч16"];
                Worksheet.Range["BE29"].Value = fieldValues["Сум9"];
                Worksheet.Range["BE30"].Value = fieldValues["Сум10"];
                Worksheet.Range["BE31"].Value = fieldValues["Сум11"];
                Worksheet.Range["BE32"].Value = fieldValues["Сум12"];
                Worksheet.Range["BE33"].Value = fieldValues["Сум13"];
                Worksheet.Range["BE34"].Value = fieldValues["Сум14"];
                Worksheet.Range["BE35"].Value = fieldValues["Сум15"];
                Worksheet.Range["BE36"].Value = fieldValues["Сум16"];

                Worksheet.Range["H38"].Value = fieldValues["Прил"];
                Worksheet.Range["T38"].Value = fieldValues["листах"];
                Worksheet.Range["W40"].Value = fieldValues["сумпропис"];
                Worksheet.Range["AM42"].Value = fieldValues["Коп"];
                Worksheet.Range["AU42"].Value = fieldValues["Руб"];
                Worksheet.Range["BF42"].Value = fieldValues["Коп"];
                Worksheet.Range["AB44"].Value = fieldValues["ГлБухРасшПодп"];
                Worksheet.Range["AB46"].Value = fieldValues["БухРасшПодп"];
                Worksheet.Range["P49"].Value = fieldValues["СУ"];
                Worksheet.Range["AB49"].Value = fieldValues["СУС"];
                Worksheet.Range["AP50"].Value = fieldValues["Н"];
                Worksheet.Range["AX50"].Value = fieldValues["ОТ"];
                Worksheet.Range["BA50"].Value = fieldValues["МЕС"];
                Worksheet.Range["BI50"].Value = fieldValues["г5"];
                Worksheet.Range["X52"].Value = fieldValues["БУХ"];
                Worksheet.Range["AX52"].Value = fieldValues["ОТ"];
                Worksheet.Range["BA52"].Value = fieldValues["МЕС"];
                Worksheet.Range["BI52"].Value = fieldValues["г5"];
                Worksheet.Range["S58"].Value = fieldValues["ФамилияИнициалы"];
                Worksheet.Range["AQ58"].Value = fieldValues["nu"];
                Worksheet.Range["AX58"].Value = fieldValues["дд"];
                Worksheet.Range["BA58"].Value = fieldValues["мм1"];
                Worksheet.Range["BI58"].Value = fieldValues["г2"];
                Worksheet.Range["L60"].Value = fieldValues["сумпропис"];
                Worksheet.Range["AL60"].Value = fieldValues["Коп"];
                Worksheet.Range["BC60"].Value = fieldValues["Прил"];
                Worksheet.Range["BG60"].Value = fieldValues["листах"];
                Worksheet.Range["Z62"].Value = fieldValues["БухРасшПодп"];
                Worksheet.Range["AX62"].Value = fieldValues["Де"];
                Worksheet.Range["BA62"].Value = fieldValues["Месяц"];
                Worksheet.Range["BI62"].Value = fieldValues["г1"];


                Worksheet = Workbook.Worksheets["Лист2"];
                Worksheet.Range["A8"].Value = fieldValues["пп1"];
                Worksheet.Range["A9"].Value = fieldValues["пп2"];
                Worksheet.Range["A10"].Value = fieldValues["пп3"];
                Worksheet.Range["A11"].Value = fieldValues["пп4"];
                Worksheet.Range["A12"].Value = fieldValues["пп5"];
                Worksheet.Range["A13"].Value = fieldValues["пп6"];
                Worksheet.Range["A14"].Value = fieldValues["пп7"];
                Worksheet.Range["A15"].Value = fieldValues["пп8"];
                Worksheet.Range["A16"].Value = fieldValues["пп9"];
                Worksheet.Range["A17"].Value = fieldValues["пп10"];

                Worksheet.Range["D8"].Value = fieldValues["д1"];
                Worksheet.Range["D9"].Value = fieldValues["д2"];
                Worksheet.Range["D10"].Value = fieldValues["д3"];
                Worksheet.Range["D11"].Value = fieldValues["д4"];
                Worksheet.Range["D12"].Value = fieldValues["д5"];
                Worksheet.Range["D13"].Value = fieldValues["д6"];
                Worksheet.Range["D14"].Value = fieldValues["д7"];
                Worksheet.Range["D15"].Value = fieldValues["д8"];
                Worksheet.Range["D16"].Value = fieldValues["д9"];
                Worksheet.Range["D17"].Value = fieldValues["д10"];

                Worksheet.Range["K8"].Value = fieldValues["нн1"];
                Worksheet.Range["K9"].Value = fieldValues["нн2"];
                Worksheet.Range["K10"].Value = fieldValues["нн3"];
                Worksheet.Range["K11"].Value = fieldValues["нн4"];
                Worksheet.Range["K12"].Value = fieldValues["нн5"];
                Worksheet.Range["K13"].Value = fieldValues["нн6"];
                Worksheet.Range["K14"].Value = fieldValues["нн7"];
                Worksheet.Range["K15"].Value = fieldValues["нн8"];
                Worksheet.Range["K16"].Value = fieldValues["нн9"];
                Worksheet.Range["K17"].Value = fieldValues["нн10"];

                Worksheet.Range["R8"].Value = fieldValues["нд1"];
                Worksheet.Range["R9"].Value = fieldValues["нд2"];
                Worksheet.Range["R10"].Value = fieldValues["нд3"];
                Worksheet.Range["R11"].Value = fieldValues["нд4"];
                Worksheet.Range["R12"].Value = fieldValues["нд5"];
                Worksheet.Range["R13"].Value = fieldValues["нд6"];
                Worksheet.Range["R14"].Value = fieldValues["нд7"];
                Worksheet.Range["R15"].Value = fieldValues["нд8"];
                Worksheet.Range["R16"].Value = fieldValues["нд9"];
                Worksheet.Range["R17"].Value = fieldValues["нд10"];

                Worksheet.Range["AC8"].Value = fieldValues["по1"];
                Worksheet.Range["AC9"].Value = fieldValues["по2"];
                Worksheet.Range["AC10"].Value = fieldValues["по3"];
                Worksheet.Range["AC11"].Value = fieldValues["по4"];
                Worksheet.Range["AC12"].Value = fieldValues["по5"];
                Worksheet.Range["AC13"].Value = fieldValues["по6"];
                Worksheet.Range["AC14"].Value = fieldValues["по7"];
                Worksheet.Range["AC15"].Value = fieldValues["по8"];
                Worksheet.Range["AC16"].Value = fieldValues["по9"];
                Worksheet.Range["AC17"].Value = fieldValues["по10"];

                Worksheet.Range["AJ8"].Value = fieldValues["пов1"];
                Worksheet.Range["AJ9"].Value = fieldValues["пов2"];
                Worksheet.Range["AJ10"].Value = fieldValues["пов3"];
                Worksheet.Range["AJ11"].Value = fieldValues["пов4"];
                Worksheet.Range["AJ12"].Value = fieldValues["пов5"];
                Worksheet.Range["AJ13"].Value = fieldValues["пов6"];
                Worksheet.Range["AJ14"].Value = fieldValues["пов7"];
                Worksheet.Range["AJ15"].Value = fieldValues["пов8"];
                Worksheet.Range["AJ16"].Value = fieldValues["пов9"];
                Worksheet.Range["AJ17"].Value = fieldValues["пов10"];

                Worksheet.Range["AQ8"].Value = fieldValues["пк1"];
                Worksheet.Range["AQ9"].Value = fieldValues["пк2"];
                Worksheet.Range["AQ10"].Value = fieldValues["пк3"];
                Worksheet.Range["AQ11"].Value = fieldValues["пк4"];
                Worksheet.Range["AQ12"].Value = fieldValues["пк5"];
                Worksheet.Range["AQ13"].Value = fieldValues["пк6"];
                Worksheet.Range["AQ14"].Value = fieldValues["пк7"];
                Worksheet.Range["AQ15"].Value = fieldValues["пк8"];
                Worksheet.Range["AQ16"].Value = fieldValues["пк9"];
                Worksheet.Range["AQ17"].Value = fieldValues["пк10"];

                Worksheet.Range["AX8"].Value = fieldValues["пкв1"];
                Worksheet.Range["AX9"].Value = fieldValues["пкв2"];
                Worksheet.Range["AX10"].Value = fieldValues["пкв3"];
                Worksheet.Range["AX11"].Value = fieldValues["пкв4"];
                Worksheet.Range["AX12"].Value = fieldValues["пкв5"];
                Worksheet.Range["AX13"].Value = fieldValues["пкв6"];
                Worksheet.Range["AX14"].Value = fieldValues["пкв7"];
                Worksheet.Range["AX15"].Value = fieldValues["пкв8"];
                Worksheet.Range["AX16"].Value = fieldValues["пкв9"];
                Worksheet.Range["AX17"].Value = fieldValues["пкв10"];

                Worksheet.Range["BE8"].Value = fieldValues["дс1"];
                Worksheet.Range["BE9"].Value = fieldValues["дс2"];
                Worksheet.Range["BE10"].Value = fieldValues["дс3"];
                Worksheet.Range["BE11"].Value = fieldValues["дс4"];
                Worksheet.Range["BE12"].Value = fieldValues["дс5"];
                Worksheet.Range["BE13"].Value = fieldValues["дс6"];
                Worksheet.Range["BE14"].Value = fieldValues["дс7"];
                Worksheet.Range["BE15"].Value = fieldValues["дс8"];
                Worksheet.Range["BE16"].Value = fieldValues["дс9"];
                Worksheet.Range["BE17"].Value = fieldValues["дс10"];

                Worksheet.Range["AC18"].Value = fieldValues["ит1"];
                Worksheet.Range["AJ18"].Value = fieldValues["ит2"];
                Worksheet.Range["AQ18"].Value = fieldValues["ит3"];
                Worksheet.Range["AX18"].Value = fieldValues["ит4"];
                Worksheet.Range["AB21"].Value = fieldValues["ФамилияИнициалы"];
                




                if (File.Exists(Environment.CurrentDirectory + @"\Report_out.xlsx"))
                {
                    File.Delete(Environment.CurrentDirectory + @"\Report_out.xlsx");
                }

                
            }
            catch(Exception ex) 
            { 
                Ex.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Ex);
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Ex.Application.ActiveWorkbook.SaveAs(Environment.CurrentDirectory + @"\Report_out.xlsx");
                Ex.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Ex);
                Process.Start(Environment.CurrentDirectory + @"\Report_out.xlsx");
            }
            

        }

        public void OpenExcelFile()
        {
            try
            {
                Ex = new Microsoft.Office.Interop.Excel.Application();
                Workbook = Ex.Workbooks.Open(Environment.CurrentDirectory + @"\Report_out.xlsx");
                Worksheet = Workbook.Worksheets["Лист1"];

                fieldValues["ТабНом"]=Worksheet.Range["BB21"].Value.ToString();
                fieldValues["Тысячи"] = Worksheet.Range["AV11"].Value.ToString();
                fieldValues["СотДесЕд"]=Worksheet.Range["AN12"].Value.ToString();
                fieldValues["Номер"]=Worksheet.Range["X14"].Value.ToString();
                fieldValues["Дата"]=Worksheet.Range["AE14"].Value.ToString();
                fieldValues["РасшПодп"]=Worksheet.Range["AX15"].Value.ToString();
                fieldValues["Де"]= Worksheet.Range["AO17"].Value.ToString();
                fieldValues["Месяц"]=Worksheet.Range["AS17"].Value.ToString();
                fieldValues["го"]=Worksheet.Range["BE17"].Value.ToString();
                fieldValues["СтруктурноеПодразделение"] =Worksheet.Range["P20"].Value.ToString();
                fieldValues["ФамилияИнициалы"] =Worksheet.Range["K21"].Value.ToString();
                fieldValues["Код"]=Worksheet.Range["BB20"].Value.ToString();
                fieldValues["проф"]= Worksheet.Range["M24"].Value.ToString();
                fieldValues["НазначениеАванса"]=Worksheet.Range["AQ24"].Value.ToString();
                fieldValues["Остаток"]= Worksheet.Range["X27"].Value.ToString();
                fieldValues["Перерасход"]=Worksheet.Range["X28"].Value.ToString();
                fieldValues["ПолАва"]=Worksheet.Range["X29"].Value.ToString();
                fieldValues["ПолВал"]=Worksheet.Range["X30"].Value.ToString();
                fieldValues["ИтогПолуч"]=Worksheet.Range["X33"].Value.ToString();
                fieldValues["Израсход"]=Worksheet.Range["X34"].Value.ToString();
                fieldValues["Остаток1"]=Worksheet.Range["X35"].Value.ToString();
                fieldValues["Перерасход1"]=Worksheet.Range["X36"].Value.ToString();
                fieldValues["Сч1"]= Worksheet.Range["AG29"].Value.ToString();
                fieldValues["Сч2"]=Worksheet.Range["AG30"].Value.ToString();
                fieldValues["Сч3"]=Worksheet.Range["AG31"].Value.ToString();
                fieldValues["Сч4"]=Worksheet.Range["AG32"].Value.ToString();
                fieldValues["Сч5"]= Worksheet.Range["AG33"].Value.ToString();
                fieldValues["Сч6"]=Worksheet.Range["AG34"].Value.ToString();
                fieldValues["Сч7"]=Worksheet.Range["AG35"].Value.ToString();
                fieldValues["Сч8"]=Worksheet.Range["AG36"].Value.ToString();
                fieldValues["Сум1"]=Worksheet.Range["AO29"].Value.ToString();
                fieldValues["Сум2"]=Worksheet.Range["AO30"].Value.ToString();
                fieldValues["Сум3"]=Worksheet.Range["AO31"].Value.ToString();
                fieldValues["Сум4"]=Worksheet.Range["AO32"].Value.ToString();
                fieldValues["Сум5"] =Worksheet.Range["AO33"].Value.ToString();
                fieldValues["Сум6"]=Worksheet.Range["AO34"].Value.ToString();
                fieldValues["Сум7"]= Worksheet.Range["AO35"].Value.ToString();
                fieldValues["Сум8"]=Worksheet.Range["AO36"].Value.ToString();
                fieldValues["Сч9"]=Worksheet.Range["AW29"].Value.ToString();
                fieldValues["Сч10"]=Worksheet.Range["AW30"].Value.ToString();
                fieldValues["Сч11"]=Worksheet.Range["AW31"].Value.ToString();
                fieldValues["Сч12"]=Worksheet.Range["AW32"].Value.ToString();
                fieldValues["Сч13"]=Worksheet.Range["AW33"].Value.ToString();
                fieldValues["Сч14"] = Worksheet.Range["AW34"].Value.ToString();
                fieldValues["Сч15"] = Worksheet.Range["AW35"].Value.ToString();
                fieldValues["Сч16"] = Worksheet.Range["AW36"].Value.ToString();
                fieldValues["Сум9"] = Worksheet.Range["BE29"].Value.ToString();
                fieldValues["Сум10"] = Worksheet.Range["BE30"].Value.ToString();
                fieldValues["Сум11"] = Worksheet.Range["BE31"].Value.ToString();
                fieldValues["Сум12"] = Worksheet.Range["BE32"].Value.ToString();
                fieldValues["Сум13"] = Worksheet.Range["BE33"].Value.ToString();
                fieldValues["Сум14"] = Worksheet.Range["BE34"].Value.ToString();
                fieldValues["Сум15"] = Worksheet.Range["BE35"].Value.ToString();
                fieldValues["Сум16"] = Worksheet.Range["BE36"].Value.ToString();
                fieldValues["Прил"] = Worksheet.Range["H38"].Value.ToString();
                fieldValues["листах"] = Worksheet.Range["T38"].Value.ToString();
                fieldValues["сумпропис"] = Worksheet.Range["W40"].Value.ToString();
                fieldValues["Коп"] = Worksheet.Range["AM42"].Value.ToString();
                fieldValues["Руб"] = Worksheet.Range["AU42"].Value.ToString();
                fieldValues["Коп"] = Worksheet.Range["BF42"].Value.ToString();
                fieldValues["ГлБухРасшПодп"] = Worksheet.Range["AB44"].Value.ToString();
                fieldValues["БухРасшПодп"] = Worksheet.Range["AB46"].Value.ToString();
                fieldValues["СУ"] = Worksheet.Range["P49"].Value.ToString();
                fieldValues["СУС"] = Worksheet.Range["AB49"].Value.ToString();
                fieldValues["Н"] = Worksheet.Range["AP50"].Value.ToString();
                fieldValues["ОТ"] = Worksheet.Range["AX50"].Value.ToString();
                fieldValues["МЕС"] = Worksheet.Range["BA50"].Value.ToString();
                fieldValues["г5"] = Worksheet.Range["BI50"].Value.ToString();
                fieldValues["БУХ"] = Worksheet.Range["X52"].Value.ToString();
                fieldValues["ОТ"] = Worksheet.Range["AX52"].Value.ToString();
                fieldValues["МЕС"] = Worksheet.Range["BA52"].Value.ToString();
                fieldValues["г5"] = Worksheet.Range["BI52"].Value.ToString();
                fieldValues["ФамилияИнициалы"] = Worksheet.Range["S58"].Value.ToString();
                fieldValues["nu"] = Worksheet.Range["AQ58"].Value.ToString();
                fieldValues["дд"] = Worksheet.Range["AX58"].Value.ToString();
                fieldValues["мм1"] = Worksheet.Range["BA58"].Value.ToString();
                fieldValues["г2"] = Worksheet.Range["BI58"].Value.ToString();
                fieldValues["сумпропис"] = Worksheet.Range["L60"].Value.ToString();
                fieldValues["Коп"] = Worksheet.Range["AL60"].Value.ToString(); 
                fieldValues["Прил"] = Worksheet.Range["BC60"].Value.ToString();
                fieldValues["листах"] = Worksheet.Range["BG60"].Value.ToString();
                fieldValues["БухРасшПодп"] = Worksheet.Range["Z62"].Value.ToString();
                fieldValues["Де"] = Worksheet.Range["AX62"].Value.ToString();
                fieldValues["Месяц"] = Worksheet.Range["BA62"].Value.ToString();
                fieldValues["г1"] = Worksheet.Range["BI62"].Value.ToString();

                Worksheet = Workbook.Worksheets["Лист2"];
                fieldValues["пп1"] = Worksheet.Range["A8"].Value.ToString();
                fieldValues["пп2"] = Worksheet.Range["A9"].Value.ToString();
                fieldValues["пп3"] = Worksheet.Range["A10"].Value.ToString();
                fieldValues["пп4"] = Worksheet.Range["A11"].Value.ToString();
                fieldValues["пп5"] = Worksheet.Range["A12"].Value.ToString();
                fieldValues["пп6"] = Worksheet.Range["A13"].Value.ToString();
                fieldValues["пп7"] = Worksheet.Range["A14"].Value.ToString();
                fieldValues["пп8"] = Worksheet.Range["A15"].Value.ToString();
                fieldValues["пп9"] = Worksheet.Range["A16"].Value.ToString();
                fieldValues["пп10"] = Worksheet.Range["A17"].Value.ToString();
                fieldValues["д1"] = Worksheet.Range["D8"].Value.ToString();
                fieldValues["д2"] = Worksheet.Range["D9"].Value.ToString();
                fieldValues["д3"] = Worksheet.Range["D10"].Value.ToString();
                fieldValues["д4"] = Worksheet.Range["D11"].Value.ToString();
                fieldValues["д5"] = Worksheet.Range["D12"].Value.ToString();
                fieldValues["д6"] = Worksheet.Range["D13"].Value.ToString();
                fieldValues["д7"] = Worksheet.Range["D14"].Value.ToString();
                fieldValues["д8"] = Worksheet.Range["D15"].Value.ToString();
                fieldValues["д9"] = Worksheet.Range["D16"].Value.ToString();
                fieldValues["д10"] = Worksheet.Range["D17"].Value.ToString();
                fieldValues["нн1"] = Worksheet.Range["K8"].Value.ToString();
                fieldValues["нн2"] = Worksheet.Range["K9"].Value.ToString();
                fieldValues["нн3"] = Worksheet.Range["K10"].Value.ToString();
                fieldValues["нн4"] = Worksheet.Range["K11"].Value.ToString();
                fieldValues["нн5"] = Worksheet.Range["K12"].Value.ToString();
                fieldValues["нн6"] = Worksheet.Range["K13"].Value.ToString();
                fieldValues["нн7"] = Worksheet.Range["K14"].Value.ToString();
                fieldValues["нн8"] = Worksheet.Range["K15"].Value.ToString();
                fieldValues["нн9"] = Worksheet.Range["K16"].Value.ToString();
                fieldValues["нн10"] = Worksheet.Range["K17"].Value.ToString();
                fieldValues["нд1"] = Worksheet.Range["R8"].Value.ToString();
                fieldValues["нд2"] = Worksheet.Range["R9"].Value.ToString();
                fieldValues["нд3"] = Worksheet.Range["R10"].Value.ToString();
                fieldValues["нд4"] = Worksheet.Range["R11"].Value.ToString();
                fieldValues["нд5"] = Worksheet.Range["R12"].Value.ToString();
                fieldValues["нд6"] = Worksheet.Range["R13"].Value.ToString();
                fieldValues["нд7"] = Worksheet.Range["R14"].Value.ToString();
                fieldValues["нд8"] = Worksheet.Range["R15"].Value.ToString();
                fieldValues["нд9"] = Worksheet.Range["R16"].Value.ToString();
                fieldValues["нд10"] = Worksheet.Range["R17"].Value.ToString();
                fieldValues["по1"] = Worksheet.Range["AC8"].Value.ToString();
                fieldValues["по2"] = Worksheet.Range["AC9"].Value.ToString();
                fieldValues["по3"] = Worksheet.Range["AC10"].Value.ToString();
                fieldValues["по4"] = Worksheet.Range["AC11"].Value.ToString();
                fieldValues["по5"] = Worksheet.Range["AC12"].Value.ToString();
                fieldValues["по6"] = Worksheet.Range["AC13"].Value.ToString();
                fieldValues["по7"] = Worksheet.Range["AC14"].Value.ToString();
                fieldValues["по8"] = Worksheet.Range["AC15"].Value.ToString();
                fieldValues["по9"] = Worksheet.Range["AC16"].Value.ToString();
                fieldValues["по10"] = Worksheet.Range["AC17"].Value.ToString();
                fieldValues["пов1"] = Worksheet.Range["AJ8"].Value.ToString();
                fieldValues["пов2"] = Worksheet.Range["AJ9"].Value.ToString();
                fieldValues["пов3"] = Worksheet.Range["AJ10"].Value.ToString();
                fieldValues["пов4"] = Worksheet.Range["AJ11"].Value.ToString();
                fieldValues["пов5"] = Worksheet.Range["AJ12"].Value.ToString();
                fieldValues["пов6"] = Worksheet.Range["AJ13"].Value.ToString();
                fieldValues["пов7"] = Worksheet.Range["AJ14"].Value.ToString();
                fieldValues["пов8"] = Worksheet.Range["AJ15"].Value.ToString();
                fieldValues["пов9"] = Worksheet.Range["AJ16"].Value.ToString();
                fieldValues["пов10"] = Worksheet.Range["AJ17"].Value.ToString();
                fieldValues["пк1"] = Worksheet.Range["AQ8"].Value.ToString();
                fieldValues["пк2"] = Worksheet.Range["AQ9"].Value.ToString();
                fieldValues["пк3"] = Worksheet.Range["AQ10"].Value.ToString();
                fieldValues["пк4"] = Worksheet.Range["AQ11"].Value.ToString();
                fieldValues["пк5"] = Worksheet.Range["AQ12"].Value.ToString();
                fieldValues["пк6"] = Worksheet.Range["AQ13"].Value.ToString();
                fieldValues["пк7"] = Worksheet.Range["AQ14"].Value.ToString();
                fieldValues["пк8"] = Worksheet.Range["AQ15"].Value.ToString();
                fieldValues["пк9"] = Worksheet.Range["AQ16"].Value.ToString();
                fieldValues["пк10"] = Worksheet.Range["AQ17"].Value.ToString();
                fieldValues["пкв1"] = Worksheet.Range["AX8"].Value.ToString();
                fieldValues["пкв2"] = Worksheet.Range["AX9"].Value.ToString();
                fieldValues["пкв3"] = Worksheet.Range["AX10"].Value.ToString();
                fieldValues["пкв4"] = Worksheet.Range["AX11"].Value.ToString();
                fieldValues["пкв5"] = Worksheet.Range["AX12"].Value.ToString();
                fieldValues["пкв6"] = Worksheet.Range["AX13"].Value.ToString();
                fieldValues["пкв7"] = Worksheet.Range["AX14"].Value.ToString();
                fieldValues["пкв8"] = Worksheet.Range["AX15"].Value.ToString();
                fieldValues["пкв9"] = Worksheet.Range["AX16"].Value.ToString();
                fieldValues["пкв10"] = Worksheet.Range["AX17"].Value.ToString();
                fieldValues["дс1"] = Worksheet.Range["BE8"].Value.ToString();
                fieldValues["дс2"] = Worksheet.Range["BE9"].Value.ToString();
                fieldValues["дс3"] = Worksheet.Range["BE10"].Value.ToString();
                fieldValues["дс4"] = Worksheet.Range["BE11"].Value.ToString();
                fieldValues["дс5"] = Worksheet.Range["BE12"].Value.ToString();
                fieldValues["дс6"] = Worksheet.Range["BE13"].Value.ToString();
                fieldValues["дс7"] = Worksheet.Range["BE14"].Value.ToString();
                fieldValues["дс8"] = Worksheet.Range["BE15"].Value.ToString();
                fieldValues["дс9"] = Worksheet.Range["BE16"].Value.ToString();
                fieldValues["дс10"] = Worksheet.Range["BE17"].Value.ToString();
                fieldValues["ит1"] = Worksheet.Range["AC18"].Value.ToString();
                fieldValues["ит2"] = Worksheet.Range["AJ18"].Value.ToString();
                fieldValues["ит3"] = Worksheet.Range["AQ18"].Value.ToString();
                fieldValues["ит4"] = Worksheet.Range["AX18"].Value.ToString();
                fieldValues["ФамилияИнициалы"] = Worksheet.Range["AB21"].Value.ToString();

                /*if (File.Exists(Environment.CurrentDirectory + @"\Report_out.xlsx"))
                {
                    File.Delete(Environment.CurrentDirectory + @"\Report_out.xlsx");
                }*/


            }
            catch(RuntimeBinderException rex)
            {

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message+"\n"+ex.StackTrace);
            }
            finally
            {
                //Ex.Application.ActiveWorkbook.SaveAs(Environment.CurrentDirectory + @"\Report_out.xlsx");
                Ex.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Ex);
                //Process.Start(Environment.CurrentDirectory + @"\Report_out.xlsx");
            }


        }

        public string GetReportName()
        {
            try
            {
                string fileName = "Отчет№" + this.fieldValues["Номер"] + "_" +
                    this.fieldValues["Дата"] + "_" + this.fieldValues["ФамилияИнициалы"] + ".docx";
                return fileName;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }
        #endregion
    }
}
