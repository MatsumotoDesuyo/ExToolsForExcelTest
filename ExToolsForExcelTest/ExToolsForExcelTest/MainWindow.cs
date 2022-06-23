using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExToolsForExcelTest
{
    public partial class MainWindow : Form
    {
        /// <summary>
        /// Excelを操作するやつ
        /// </summary>
        public ExcelController excelController { get; private set; }
        /// <summary>
        /// ホットキーの設定をするやつ
        /// </summary>
        public HotKeyController hotKeyController { get; private set; }
        /// <summary>
        /// シングルトン
        /// </summary>
        public static MainWindow Instance { get; private set; }

        public MainWindow()
        {
            Instance = this;
            InitializeComponent();
            this.ShowInTaskbar = false;
            exitStripMenuItem.Click += ExitMenuItem_Click;
        }

        private void MainWindow_Load(object sender, EventArgs e)
        {
            excelController=new ExcelController();
            hotKeyController = new HotKeyController();
            initializeSettings();
        }

        private void ExitMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            e.Cancel = true;
            this.Visible = false;
            initializeSettings();
        }

        private void notifyIcon_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (WindowState == FormWindowState.Minimized)
                {
                    WindowState = FormWindowState.Normal;
                    Visible = true;
                }
                else
                {
                    Visible = !Visible;
                }
                if (Visible) { Activate(); }
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);
            hotKeyController.ClearHotKeys();
        }

        

        void initializeSettings()
        {
            workbookNameTextBox.Text = excelController.WorkbookKey;
            testSheetNameTextBox.Text = excelController.TestSheetKey;
            evidenceSheetTextBox.Text = excelController.EvidenceSheetKey;
            passedTextBox.Text = excelController.PassedText;
            ngTextBox.Text = excelController.FailureText;
            numberColumnNameTextBox.Text = excelController.TestNumberColumn;
            descriptionColumnNameTextBox.Text = excelController.TestDescriptionColumn;
            resultColumnNameTextBox.Text = excelController.TestResultColumn;
            marginTopTextBox.Text = excelController.MarginTop.ToString();
            imageMagnificationTextBox.Text = excelController.ImageMagnification.ToString();
        }

        private void saveButton_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(workbookNameTextBox.Text))
            {
                excelController.WorkbookKey = workbookNameTextBox.Text;
            }
            if (!string.IsNullOrEmpty(testSheetNameTextBox.Text))
            {
                excelController.TestSheetKey = testSheetNameTextBox.Text;
            }
            if (!string.IsNullOrEmpty(evidenceSheetTextBox.Text))
            {
                excelController.EvidenceSheetKey = evidenceSheetTextBox.Text;
            }
            excelController.PassedText = passedTextBox.Text;
            excelController.FailureText = ngTextBox.Text;
            if (Regex.IsMatch(numberColumnNameTextBox.Text, @"^[a-zA-Z]+$"))
            {
                excelController.TestNumberColumn = numberColumnNameTextBox.Text;
            }
            if (Regex.IsMatch(descriptionColumnNameTextBox.Text, @"^[a-zA-Z]+$"))
            {
                excelController.TestDescriptionColumn = descriptionColumnNameTextBox.Text;
            }
            if (Regex.IsMatch(resultColumnNameTextBox.Text, @"^[a-zA-Z]+$"))
            {
                excelController.TestResultColumn = resultColumnNameTextBox.Text;
            }
            int marginTop;
            if (int.TryParse(marginTopTextBox.Text, out marginTop))
            {
                excelController.MarginTop = marginTop;
            }
            decimal imageMagnification;
            if(decimal.TryParse(imageMagnificationTextBox.Text,out imageMagnification))
            {
                excelController.ImageMagnification = imageMagnification;
            }

            excelController.SaveSettings();
            initializeSettings();
        }

        private void initializeButton_Click(object sender, EventArgs e)
        {
            initializeSettings();
        }

        private void testPassedShortcutTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            e.Handled = true;
            testPassedShortcutKey.SetKey(e.KeyCode);
            testPassedShortcutTextBox.Text = testPassedShortcutKey.ToString();
        }

        ShortcutKey testPassedShortcutKey = new ShortcutKey();

        private void testPassedShortcutTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void testPassedShortcutTextBox_KeyUp(object sender, KeyEventArgs e)
        {
            e.Handled = true;
            if (testPassedShortcutKey.Key == Keys.None)
            {
                testPassedShortcutTextBox.Text = "";
            }
        }
    }
}
