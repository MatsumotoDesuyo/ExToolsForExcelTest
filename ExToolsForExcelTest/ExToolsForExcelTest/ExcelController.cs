﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Range=Microsoft.Office.Interop.Excel.Range;

namespace ExToolsForExcelTest
{
    public class ExcelController
    {
        public Microsoft.Office.Interop.Excel.Application Excel = null;
        public Microsoft.Office.Interop.Excel.Workbook Workbook = null;
        public Microsoft.Office.Interop.Excel.Worksheet TestSheet = null;
        public Microsoft.Office.Interop.Excel.Worksheet EvidenceSheet = null;

        public string WorkbookKey = "動作検証仕様書";
        public string TestSheetKey = "動作検証仕様書";
        public string EvidenceSheetKey = "エビデンス";

        public string PassedText = "〇";
        public string FailureText = "×";

        public string TestNumberColumn = "A";
        public string TestDescriptionColumn = "M";
        public string TestResultColumn = "AF";

        public int MarginTop = 2;
        public decimal ImageMagnification = 1;

        const string CaptureImg = "Capture.bmp";

        public ExcelController()
        {
            LoadSettings();
            setExcelObjects();
            var processes = Process.GetProcesses();
            foreach (var process in processes)
            {
                var name = process.ProcessName;
                Console.WriteLine(name);
            }
        }

        public void WriteTestResult(bool correct)
        {
            if (Excel == null || Workbook == null || TestSheet == null)
            {
                setExcelObjects();
                if (Excel == null || Workbook == null || TestSheet == null)
                {
                    return;
                }
            }

            Range activeCell = Excel.ActiveCell;
            int row = activeCell.Row;
            TestSheet.Range[TestResultColumn + row + ":" + 
                (TestResultColumn.ToColumnNumber()+1).ToColumnName()+ row].Value = 
                correct?PassedText:FailureText;
            Excel.Goto(TestSheet.Range["A" + (row + 1 - MarginTop)], true);
            TestSheet.Range[TestResultColumn + (row + 1)].Activate();
        }
        public void WriteTestResultWithEvidence(bool correct)
        {
            if (Excel == null || Workbook == null || TestSheet == null||EvidenceSheet==null)
            {
                setExcelObjects();
                if (Excel == null || Workbook == null || TestSheet == null)
                {
                    return;
                }
            }

            Range activeCell = Excel.ActiveCell;
            int row = activeCell.Row;
            TestSheet.Range[TestResultColumn + row + ":" +
                (TestResultColumn.ToColumnNumber() + 1).ToColumnName() + row].Value =
                correct ? PassedText : FailureText;
            Excel.Goto(TestSheet.Range["A" + (row + 1 - MarginTop)], true);
            TestSheet.Range[TestResultColumn + (row + 1)].Activate();

            if (EvidenceSheet == null) { return; }
            object obj=TestSheet.Range[TestNumberColumn + row].Value;
            int testNumber;
            if(!(int.TryParse(obj.ToString(),out testNumber)))
            {
                return;
            }
            if (testNumber <= 0) { return; }

            EvidenceSheet.Range["A" + (testNumber * 2 - 1)].Value = testNumber;
            EvidenceSheet.Range["B" + (testNumber * 2 - 1)].Value = TestSheet.Range[TestDescriptionColumn + row + ":" + TestDescriptionColumn.NextColumnName() + row].Value;

            Bitmap bitmap = WindowCapture.CaptureActiveWindow();
            WindowCapture.SaveActiveWindowCapture(bitmap,CaptureImg);
            Range captureRange = EvidenceSheet.Range["B" + (testNumber * 2)];
            var shape = EvidenceSheet.Shapes.AddPicture(
                System.IO.Directory.GetCurrentDirectory()+@"\"+CaptureImg, 
                Microsoft.Office.Core.MsoTriState.msoCTrue, 
                Microsoft.Office.Core.MsoTriState.msoCTrue, 
                captureRange.Left, captureRange.Top, 
                (float)ImageMagnification*bitmap.Width,
                (float)ImageMagnification*bitmap.Height);
            EvidenceSheet.Rows[testNumber * 2].RowHeight = shape.Height;
        }
        public void SkipRow()
        {
            if (Excel == null || Workbook == null || TestSheet == null)
            {
                setExcelObjects();
                if (Excel == null || Workbook == null || TestSheet == null)
                {
                    return;
                }
            }

            Range activeCell = Excel.ActiveCell;
            int row = activeCell.Row;
            Excel.Goto(TestSheet.Range["A" + (row + 1 - MarginTop)], true);
            TestSheet.Range[TestResultColumn + (row + 1)].Activate();
        }

        public bool SaveSettings()
        {
            ExcelControllerSettings settings = new ExcelControllerSettings()
            {
                WorkbookKey = this.WorkbookKey,
                TestSheetKey = this.TestSheetKey,
                EvidenceSheetKey = this.EvidenceSheetKey,
                PassedText = this.PassedText,
                FailureText = this.FailureText,
                TestNumberColumn = this.TestNumberColumn,
                TestDescriptionColumn = this.TestDescriptionColumn,
                TestResultColumn = this.TestResultColumn,
                MarginTop = this.MarginTop,
                ImageMagnification = this.ImageMagnification
            };
            return settings.Save();
        }
        public bool LoadSettings()
        {
            ExcelControllerSettings settings = ExcelControllerSettings.Load();
            WorkbookKey = settings.WorkbookKey;
            TestSheetKey = settings.TestSheetKey;
            EvidenceSheetKey = settings.EvidenceSheetKey;
            PassedText = settings.PassedText;
            FailureText = settings.FailureText;
            TestNumberColumn = settings.TestNumberColumn;
            TestDescriptionColumn = settings.TestDescriptionColumn;
            TestResultColumn = settings.TestResultColumn;
            MarginTop = settings.MarginTop;
            ImageMagnification = settings.ImageMagnification;
            return true;
        }

        void setExcelObjects()
        {
            clearExcelObjects();
            try
            {
                object obj = Marshal.GetActiveObject("Excel.Application");
                if (obj == null) return;
                Excel = (Microsoft.Office.Interop.Excel.Application)obj;
            }
            catch { return; }
            foreach (Microsoft.Office.Interop.Excel.Workbook wb in Excel.Workbooks)
            {
                if (wb.Name.Contains(WorkbookKey))
                {
                    Workbook = wb;
                    foreach (Microsoft.Office.Interop.Excel.Worksheet st in wb.Worksheets)
                    {
                        if (st.Name.Contains(TestSheetKey))
                        {
                            TestSheet = st;
                            
                        }
                        else if (st.Name.Contains(EvidenceSheetKey))
                        {
                            EvidenceSheet = st;
                        }
                    }
                    break;
                }
            }

            if (Workbook != null)
            {
                Workbook.BeforeClose += (ref bool cancel) => clearExcelObjects();
            }
            if (TestSheet != null)
            {
                TestSheet.BeforeDelete += () => clearExcelObjects();
            }
            if (EvidenceSheet != null)
            {
                EvidenceSheet.BeforeDelete += () => clearExcelObjects();
            }
        }
        void clearExcelObjects()
        {
            if (Excel != null)
            {
                Marshal.ReleaseComObject(Excel);
            }
            Excel = null;
            if (Workbook != null)
            {
                Marshal.ReleaseComObject(Workbook);
            }
            Workbook = null;
            if (TestSheet != null)
            {
                Marshal.ReleaseComObject(TestSheet);
            }
            TestSheet = null;
            if (EvidenceSheet != null)
            {
                Marshal.ReleaseComObject(EvidenceSheet);
            }
            EvidenceSheet = null;
        }

        ~ExcelController()
        {
            clearExcelObjects();
        }



        public static int ToColumnNumber(string source)
        {
            return source.ToColumnNumber();
        }
        public static string ToColumnName(int source)
        {
            return source.ToColumnName();
        }
    }

    class ExcelControllerSettings
    {
        public string WorkbookKey { get; set; }
        public string TestSheetKey { get; set; }
        public string EvidenceSheetKey { get; set; }

        public string PassedText { get; set; }
        public string FailureText { get; set; }

        public string TestNumberColumn { get; set; }
        public string TestDescriptionColumn { get; set; }
        public string TestResultColumn { get; set; }

        public int MarginTop { get; set; }
        public decimal ImageMagnification { get; set; }

        public ExcelControllerSettings() { }

        public bool Save()
        {
            string jsonString = JsonSerializer.Serialize(this);
            try
            {
                File.WriteAllText(ConfigurationManager.AppSettings[ConfigKeys.ExcelSettingFilePath], jsonString);
            }
            catch
            {
                return false;
            }
            return true;
        }
        public static ExcelControllerSettings Load(string path)
        {
            try
            {
                string jsonString=File.ReadAllText(path);
                return JsonSerializer.Deserialize<ExcelControllerSettings>(jsonString);
            }
            catch
            {
                return null;
            }
        }
        public static ExcelControllerSettings Load()
        {
            return Load(ConfigurationManager.AppSettings[ConfigKeys.ExcelSettingFilePath]);
        }
    }

    static class ExcelControllerExtend
    {
        public static int ToColumnNumber(this string source)
        {
            if (string.IsNullOrWhiteSpace(source))
                return 0;

            string buf = source.ToUpper();
            if (Regex.IsMatch(buf, @"^[A-Z]+$") == false)
                throw new FormatException("Argument format is only alphabet");

            // 変換後がint.MaxValueに収まるか?
            if (buf.Length>=7&&buf.CompareTo("FXSHRXW") >= 1)
                throw new ArgumentOutOfRangeException("Argument range max \"FXSHRXW\"");

            return _ToColumnNumber(buf, 0);
        }
        public static string ToColumnName(this int source)
        {
            if (source < 1) return string.Empty;
            return ToColumnName((source - 1) / 26) + (char)('A' + ((source - 1) % 26));
        }
        static int _ToColumnNumber(this string source, int call_num)
        {
            if (source == "") return 0;
            int digit = (int)Math.Pow(26, call_num);
            return ((source.Last() - 'A' + 1) * digit) + _ToColumnNumber(source.Substring(0, source.Length - 1), ++call_num);
        }
        public static string NextColumnName(this string source)
        {
            return (source.ToColumnNumber() + 1).ToColumnName();
        }
    }
}
