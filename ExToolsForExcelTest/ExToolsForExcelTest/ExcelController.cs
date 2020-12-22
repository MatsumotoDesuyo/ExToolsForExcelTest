using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
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

        const string CaptureImg = "Capture.bmp";

        public ExcelController()
        {
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

            WindowCapture.SaveActiveWindowCapture(CaptureImg);
            Range captureRange = EvidenceSheet.Range["B" + (testNumber * 2)];
            var shape = EvidenceSheet.Shapes.AddPicture(System.IO.Directory.GetCurrentDirectory()+@"\"+CaptureImg, Microsoft.Office.Core.MsoTriState.msoCTrue, Microsoft.Office.Core.MsoTriState.msoCTrue, captureRange.Left, captureRange.Top, 100, 100);
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

        void setExcelObjects()
        {
            clearExcelObjects();
            object obj = Marshal.GetActiveObject("Excel.Application");
            if (obj == null) return;
            Excel = (Microsoft.Office.Interop.Excel.Application)obj;
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
