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

namespace ExToolsForExcelTest
{
    public partial class MainWindow : Form
    {
        ExcelController excelController;
        HotKey writeTestOkHotKey;
        HotKey writeTestNgHotKey;
        HotKey skipRowHotKey;
        HotKey writeTestOkWithEvidenceHotKey;
        HotKey writeTestNgWithEvidenceHotKey;

        public MainWindow()
        {
            InitializeComponent();
            this.ShowInTaskbar = false;
            exitStripMenuItem.Click += ExitMenuItem_Click;
        }

        private void MainWindow_Load(object sender, EventArgs e)
        {
            excelController=new ExcelController();
            setHotKeys();
        }

        private void ExitMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            e.Cancel = true;
            this.Visible = false;
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

        void setHotKeys()
        {
            writeTestOkHotKey = new HotKey(MOD_KEY.ALT, Keys.D9);
            writeTestOkHotKey.HotKeyPush += new EventHandler((obj, args) => { excelController.WriteTestResult(true); });
            writeTestNgHotKey = new HotKey(MOD_KEY.ALT|MOD_KEY.CONTROL, Keys.D9);
            writeTestNgHotKey.HotKeyPush += new EventHandler((obj, args) => { excelController.WriteTestResult(false); });
            skipRowHotKey = new HotKey(MOD_KEY.ALT, Keys.D0);
            skipRowHotKey.HotKeyPush += new EventHandler((obj, args) => { excelController.SkipRow(); });
            writeTestOkWithEvidenceHotKey = new HotKey(MOD_KEY.ALT, Keys.D8);
            writeTestOkWithEvidenceHotKey.HotKeyPush += new EventHandler((obj, args) => { excelController.WriteTestResultWithEvidence(true); });
            writeTestNgWithEvidenceHotKey = new HotKey(MOD_KEY.ALT | MOD_KEY.CONTROL, Keys.D8);
            writeTestNgWithEvidenceHotKey.HotKeyPush += new EventHandler((obj, args) => { excelController.WriteTestResultWithEvidence(false); });
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);
            clearHotKeys();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            excelController.WriteTestResult(true);
        }

        void clearHotKeys()
        {
            if(writeTestOkHotKey!=null)
                writeTestOkHotKey.Dispose();
            if(writeTestNgHotKey!=null)
                writeTestNgHotKey.Dispose();
            if(skipRowHotKey!=null)
                skipRowHotKey.Dispose();
            if (writeTestOkWithEvidenceHotKey != null)
                writeTestOkWithEvidenceHotKey.Dispose();
            if (writeTestNgWithEvidenceHotKey != null)
            {
                writeTestNgWithEvidenceHotKey.Dispose();
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            Console.WriteLine(e.KeyCode);
        }
    }
}
