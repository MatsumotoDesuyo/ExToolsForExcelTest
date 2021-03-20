using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExToolsForExcelTest
{
    public class HotKeyController
    {
        public ShortcutKey WriteTestOkShortcut;
        public ShortcutKey WriteTestNgShortcut;
        public ShortcutKey SkipRowShortcut;
        public ShortcutKey WriteTestOkWithEvidenceShortcut;
        public ShortcutKey WriteTestNgWithEvidenceShortcut;

        HotKey writeTestOkHotKey;
        HotKey writeTestNgHotKey;
        HotKey skipRowHotKey;
        HotKey writeTestOkWithEvidenceHotKey;
        HotKey writeTestNgWithEvidenceHotKey;

        public HotKeyController()
        {
            SetHotKeys();
        }

        public void SetHotKeys()
        {
            writeTestOkHotKey = new HotKey(MOD_KEY.ALT, Keys.D9);
            writeTestOkHotKey.HotKeyPush += new EventHandler((obj, args) => { MainWindow.Instance.excelController.WriteTestResult(true); });
            writeTestNgHotKey = new HotKey(MOD_KEY.ALT | MOD_KEY.CONTROL, Keys.D9);
            writeTestNgHotKey.HotKeyPush += new EventHandler((obj, args) => { MainWindow.Instance.excelController.WriteTestResult(false); });
            skipRowHotKey = new HotKey(MOD_KEY.ALT, Keys.D0);
            skipRowHotKey.HotKeyPush += new EventHandler((obj, args) => { MainWindow.Instance.excelController.SkipRow(); });
            writeTestOkWithEvidenceHotKey = new HotKey(MOD_KEY.ALT, Keys.D8);
            writeTestOkWithEvidenceHotKey.HotKeyPush += new EventHandler((obj, args) => { MainWindow.Instance.excelController.WriteTestResultWithEvidence(true); });
            writeTestNgWithEvidenceHotKey = new HotKey(MOD_KEY.ALT | MOD_KEY.CONTROL, Keys.D8);
            writeTestNgWithEvidenceHotKey.HotKeyPush += new EventHandler((obj, args) => { MainWindow.Instance.excelController.WriteTestResultWithEvidence(false); });
        }

        public void ClearHotKeys()
        {
            if (writeTestOkHotKey != null)
                writeTestOkHotKey.Dispose();
            if (writeTestNgHotKey != null)
                writeTestNgHotKey.Dispose();
            if (skipRowHotKey != null)
                skipRowHotKey.Dispose();
            if (writeTestOkWithEvidenceHotKey != null)
                writeTestOkWithEvidenceHotKey.Dispose();
            if (writeTestNgWithEvidenceHotKey != null)
            {
                writeTestNgWithEvidenceHotKey.Dispose();
            }
        }
    
    }

    class ShortcutKeySettings
    {
        public ShortcutKey WriteTestOkShortcut;
        public ShortcutKey WriteTestNgShortcut;
        public ShortcutKey SkipRowShortcut;
        public ShortcutKey WriteTestOkWithEvidenceShortcut;
        public ShortcutKey WriteTestNgWithEvidenceShortcut;

        public ShortcutKeySettings() { }

        public bool Save()
        {
            string jsonString = JsonSerializer.Serialize(this);
            try
            {
                File.WriteAllText(ConfigurationManager.AppSettings[ConfigKeys.HotKeySettingFilePath], jsonString);
            }
            catch
            {
                return false;
            }
            return true;
        }
        public static ShortcutKeySettings Load(string path)
        {
            try
            {
                string jsonString = File.ReadAllText(path);
                return JsonSerializer.Deserialize<ShortcutKeySettings>(jsonString);
            }
            catch
            {
                return null;
            }
        }
        public static ShortcutKeySettings Load()
        {
            return Load(ConfigurationManager.AppSettings[ConfigKeys.HotKeySettingFilePath]);
        }
    }

    public class ShortcutKey
    {
        public bool Ctrl;
        public bool Alt;
        public bool Shift;
        public Keys Key = Keys.None;

        public void SetKey(Keys keys)
        {
            Ctrl = (Control.ModifierKeys & Keys.Control) == Keys.Control;
            Alt=(Control.ModifierKeys& Keys.Alt) == Keys.Alt;
            Shift = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;

            if (!Ctrl && !Alt && !Shift) {
                Key = Keys.None;
                return; 
            }
            if (keys!=Keys.ShiftKey&&keys!=Keys.ControlKey&&keys!=Keys.Menu)
            {
                Key = keys;
            }
            else
            {
                Key = Keys.None;
            }
        }

        public override string ToString()
        {
            List<string> strs = new List<string>();
            if (Ctrl)
            {
                strs.Add("Ctrl");
            }
            if (Alt)
            {
                strs.Add("Alt");
            }
            if (Shift)
            {
                strs.Add("Shift");
            }
            if (Key != Keys.None)
            {
                strs.Add(Key.ToString());
            }
            string str = "";
            for(int i = 0; i < strs.Count; i++)
            {
                if (i != 0)
                {
                    str += " + ";
                }
                str += strs[i];
            }
            return str;
        }
    }
}
