
namespace ExToolsForExcelTest
{
    partial class MainWindow
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainWindow));
            this.notifyIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.contextMenuStrip = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.exitStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveButton = new System.Windows.Forms.Button();
            this.initializeButton = new System.Windows.Forms.Button();
            this.excelSettingsLabelLayoutPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.excelSettingsLabel = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label16 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.imageMagnificationTextBox = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.marginTopTextBox = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.resultColumnNameTextBox = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.descriptionColumnNameTextBox = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.numberColumnNameTextBox = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.ngTextBox = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.passedTextBox = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.evidenceSheetTextBox = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.testSheetNameTextBox = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.workbookKeyLabel = new System.Windows.Forms.Label();
            this.workbookNameTextBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label18 = new System.Windows.Forms.Label();
            this.testPassedShortcutTextBox = new System.Windows.Forms.TextBox();
            this.label19 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label20 = new System.Windows.Forms.Label();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.label21 = new System.Windows.Forms.Label();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.label22 = new System.Windows.Forms.Label();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.contextMenuStrip.SuspendLayout();
            this.excelSettingsLabelLayoutPanel.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // notifyIcon
            // 
            this.notifyIcon.ContextMenuStrip = this.contextMenuStrip;
            this.notifyIcon.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon.Icon")));
            this.notifyIcon.Text = "notifyIcon";
            this.notifyIcon.Visible = true;
            this.notifyIcon.MouseClick += new System.Windows.Forms.MouseEventHandler(this.notifyIcon_MouseClick);
            // 
            // contextMenuStrip
            // 
            this.contextMenuStrip.ImageScalingSize = new System.Drawing.Size(32, 32);
            this.contextMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exitStripMenuItem});
            this.contextMenuStrip.Name = "contextMenuStrip";
            this.contextMenuStrip.Size = new System.Drawing.Size(139, 42);
            // 
            // exitStripMenuItem
            // 
            this.exitStripMenuItem.Name = "exitStripMenuItem";
            this.exitStripMenuItem.Size = new System.Drawing.Size(138, 38);
            this.exitStripMenuItem.Text = "&終了";
            // 
            // saveButton
            // 
            this.saveButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.saveButton.Location = new System.Drawing.Point(663, 787);
            this.saveButton.Name = "saveButton";
            this.saveButton.Size = new System.Drawing.Size(151, 59);
            this.saveButton.TabIndex = 15;
            this.saveButton.Text = "適用(&O)";
            this.saveButton.UseVisualStyleBackColor = true;
            this.saveButton.Click += new System.EventHandler(this.saveButton_Click);
            // 
            // initializeButton
            // 
            this.initializeButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.initializeButton.Location = new System.Drawing.Point(831, 787);
            this.initializeButton.Name = "initializeButton";
            this.initializeButton.Size = new System.Drawing.Size(151, 59);
            this.initializeButton.TabIndex = 14;
            this.initializeButton.Text = "元に戻す(&C)";
            this.initializeButton.UseVisualStyleBackColor = true;
            this.initializeButton.Click += new System.EventHandler(this.initializeButton_Click);
            // 
            // excelSettingsLabelLayoutPanel
            // 
            this.excelSettingsLabelLayoutPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.excelSettingsLabelLayoutPanel.Controls.Add(this.excelSettingsLabel);
            this.excelSettingsLabelLayoutPanel.FlowDirection = System.Windows.Forms.FlowDirection.BottomUp;
            this.excelSettingsLabelLayoutPanel.Location = new System.Drawing.Point(2, 9);
            this.excelSettingsLabelLayoutPanel.Margin = new System.Windows.Forms.Padding(0);
            this.excelSettingsLabelLayoutPanel.Name = "excelSettingsLabelLayoutPanel";
            this.excelSettingsLabelLayoutPanel.Padding = new System.Windows.Forms.Padding(15, 0, 0, 0);
            this.excelSettingsLabelLayoutPanel.Size = new System.Drawing.Size(980, 41);
            this.excelSettingsLabelLayoutPanel.TabIndex = 16;
            // 
            // excelSettingsLabel
            // 
            this.excelSettingsLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.excelSettingsLabel.AutoSize = true;
            this.excelSettingsLabel.Location = new System.Drawing.Point(18, 17);
            this.excelSettingsLabel.Name = "excelSettingsLabel";
            this.excelSettingsLabel.Size = new System.Drawing.Size(112, 24);
            this.excelSettingsLabel.TabIndex = 0;
            this.excelSettingsLabel.Text = "Excel設定";
            // 
            // panel2
            // 
            this.panel2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel2.AutoSize = true;
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.label16);
            this.panel2.Controls.Add(this.label15);
            this.panel2.Controls.Add(this.imageMagnificationTextBox);
            this.panel2.Controls.Add(this.label14);
            this.panel2.Controls.Add(this.marginTopTextBox);
            this.panel2.Controls.Add(this.label13);
            this.panel2.Controls.Add(this.resultColumnNameTextBox);
            this.panel2.Controls.Add(this.label12);
            this.panel2.Controls.Add(this.descriptionColumnNameTextBox);
            this.panel2.Controls.Add(this.label11);
            this.panel2.Controls.Add(this.numberColumnNameTextBox);
            this.panel2.Controls.Add(this.label9);
            this.panel2.Controls.Add(this.ngTextBox);
            this.panel2.Controls.Add(this.label10);
            this.panel2.Controls.Add(this.label7);
            this.panel2.Controls.Add(this.passedTextBox);
            this.panel2.Controls.Add(this.label8);
            this.panel2.Controls.Add(this.label5);
            this.panel2.Controls.Add(this.evidenceSheetTextBox);
            this.panel2.Controls.Add(this.label6);
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.testSheetNameTextBox);
            this.panel2.Controls.Add(this.label4);
            this.panel2.Controls.Add(this.label2);
            this.panel2.Controls.Add(this.workbookKeyLabel);
            this.panel2.Controls.Add(this.workbookNameTextBox);
            this.panel2.Controls.Add(this.label1);
            this.panel2.Location = new System.Drawing.Point(17, 56);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(955, 402);
            this.panel2.TabIndex = 21;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(331, 343);
            this.label16.Margin = new System.Windows.Forms.Padding(0);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(34, 24);
            this.label16.TabIndex = 47;
            this.label16.Text = "倍";
            this.label16.TextAlign = System.Drawing.ContentAlignment.BottomRight;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(47, 340);
            this.label15.Margin = new System.Windows.Forms.Padding(0);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(206, 24);
            this.label15.TabIndex = 46;
            this.label15.Text = "エビデンス画像サイズ";
            this.label15.TextAlign = System.Drawing.ContentAlignment.BottomRight;
            // 
            // imageMagnificationTextBox
            // 
            this.imageMagnificationTextBox.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.imageMagnificationTextBox.Location = new System.Drawing.Point(256, 340);
            this.imageMagnificationTextBox.Name = "imageMagnificationTextBox";
            this.imageMagnificationTextBox.Size = new System.Drawing.Size(70, 31);
            this.imageMagnificationTextBox.TabIndex = 45;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(123, 303);
            this.label14.Margin = new System.Windows.Forms.Padding(0);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(130, 24);
            this.label14.TabIndex = 44;
            this.label14.Text = "上余白列数";
            this.label14.TextAlign = System.Drawing.ContentAlignment.BottomRight;
            // 
            // marginTopTextBox
            // 
            this.marginTopTextBox.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.marginTopTextBox.Location = new System.Drawing.Point(256, 303);
            this.marginTopTextBox.Name = "marginTopTextBox";
            this.marginTopTextBox.Size = new System.Drawing.Size(70, 31);
            this.marginTopTextBox.TabIndex = 43;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(68, 268);
            this.label13.Margin = new System.Windows.Forms.Padding(0);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(185, 24);
            this.label13.TabIndex = 42;
            this.label13.Text = "テスト結果のカラム";
            this.label13.TextAlign = System.Drawing.ContentAlignment.BottomRight;
            // 
            // resultColumnNameTextBox
            // 
            this.resultColumnNameTextBox.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.resultColumnNameTextBox.Location = new System.Drawing.Point(256, 265);
            this.resultColumnNameTextBox.Name = "resultColumnNameTextBox";
            this.resultColumnNameTextBox.Size = new System.Drawing.Size(70, 31);
            this.resultColumnNameTextBox.TabIndex = 41;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(68, 231);
            this.label12.Margin = new System.Windows.Forms.Padding(0);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(185, 24);
            this.label12.TabIndex = 40;
            this.label12.Text = "テスト内容のカラム";
            this.label12.TextAlign = System.Drawing.ContentAlignment.BottomRight;
            // 
            // descriptionColumnNameTextBox
            // 
            this.descriptionColumnNameTextBox.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.descriptionColumnNameTextBox.Location = new System.Drawing.Point(256, 228);
            this.descriptionColumnNameTextBox.Name = "descriptionColumnNameTextBox";
            this.descriptionColumnNameTextBox.Size = new System.Drawing.Size(70, 31);
            this.descriptionColumnNameTextBox.TabIndex = 39;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(68, 195);
            this.label11.Margin = new System.Windows.Forms.Padding(0);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(185, 24);
            this.label11.TabIndex = 38;
            this.label11.Text = "テスト番号のカラム";
            this.label11.TextAlign = System.Drawing.ContentAlignment.BottomRight;
            // 
            // numberColumnNameTextBox
            // 
            this.numberColumnNameTextBox.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.numberColumnNameTextBox.Location = new System.Drawing.Point(256, 192);
            this.numberColumnNameTextBox.Name = "numberColumnNameTextBox";
            this.numberColumnNameTextBox.Size = new System.Drawing.Size(70, 31);
            this.numberColumnNameTextBox.TabIndex = 36;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(47, 158);
            this.label9.Margin = new System.Windows.Forms.Padding(0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(208, 24);
            this.label9.TabIndex = 35;
            this.label9.Text = "テスト不可時テキスト";
            this.label9.TextAlign = System.Drawing.ContentAlignment.BottomRight;
            // 
            // ngTextBox
            // 
            this.ngTextBox.Location = new System.Drawing.Point(256, 155);
            this.ngTextBox.Name = "ngTextBox";
            this.ngTextBox.Size = new System.Drawing.Size(482, 31);
            this.ngTextBox.TabIndex = 33;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(744, 158);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(76, 24);
            this.label10.TabIndex = 34;
            this.label10.Text = "を挿入";
            this.label10.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(47, 121);
            this.label7.Margin = new System.Windows.Forms.Padding(0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(208, 24);
            this.label7.TabIndex = 32;
            this.label7.Text = "テスト通過時テキスト";
            this.label7.TextAlign = System.Drawing.ContentAlignment.BottomRight;
            // 
            // passedTextBox
            // 
            this.passedTextBox.Location = new System.Drawing.Point(256, 118);
            this.passedTextBox.Name = "passedTextBox";
            this.passedTextBox.Size = new System.Drawing.Size(482, 31);
            this.passedTextBox.TabIndex = 30;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(744, 121);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(76, 24);
            this.label8.TabIndex = 31;
            this.label8.Text = "を挿入";
            this.label8.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(47, 84);
            this.label5.Margin = new System.Windows.Forms.Padding(0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(206, 24);
            this.label5.TabIndex = 29;
            this.label5.Text = "エビデンス用シート名";
            this.label5.TextAlign = System.Drawing.ContentAlignment.BottomRight;
            // 
            // evidenceSheetTextBox
            // 
            this.evidenceSheetTextBox.Location = new System.Drawing.Point(256, 81);
            this.evidenceSheetTextBox.Name = "evidenceSheetTextBox";
            this.evidenceSheetTextBox.Size = new System.Drawing.Size(482, 31);
            this.evidenceSheetTextBox.TabIndex = 27;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(744, 84);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(72, 24);
            this.label6.TabIndex = 28;
            this.label6.Text = "を含む";
            this.label6.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(85, 47);
            this.label3.Margin = new System.Windows.Forms.Padding(0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(168, 24);
            this.label3.TabIndex = 26;
            this.label3.Text = "テスト用シート名";
            this.label3.TextAlign = System.Drawing.ContentAlignment.BottomRight;
            // 
            // testSheetNameTextBox
            // 
            this.testSheetNameTextBox.Location = new System.Drawing.Point(256, 44);
            this.testSheetNameTextBox.Name = "testSheetNameTextBox";
            this.testSheetNameTextBox.Size = new System.Drawing.Size(482, 31);
            this.testSheetNameTextBox.TabIndex = 24;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(744, 47);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(72, 24);
            this.label4.TabIndex = 25;
            this.label4.Text = "を含む";
            this.label4.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(115, 10);
            this.label2.Margin = new System.Windows.Forms.Padding(0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(138, 24);
            this.label2.TabIndex = 23;
            this.label2.Text = "ワークブック名";
            this.label2.TextAlign = System.Drawing.ContentAlignment.BottomRight;
            // 
            // workbookKeyLabel
            // 
            this.workbookKeyLabel.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.workbookKeyLabel.Location = new System.Drawing.Point(-272, 190);
            this.workbookKeyLabel.Name = "workbookKeyLabel";
            this.workbookKeyLabel.Size = new System.Drawing.Size(179, 23);
            this.workbookKeyLabel.TabIndex = 20;
            this.workbookKeyLabel.Text = "ワークブック名";
            this.workbookKeyLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // workbookNameTextBox
            // 
            this.workbookNameTextBox.Location = new System.Drawing.Point(256, 7);
            this.workbookNameTextBox.Name = "workbookNameTextBox";
            this.workbookNameTextBox.Size = new System.Drawing.Size(482, 31);
            this.workbookNameTextBox.TabIndex = 21;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(744, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(72, 24);
            this.label1.TabIndex = 22;
            this.label1.Text = "を含む";
            this.label1.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(20, 487);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(127, 24);
            this.label17.TabIndex = 48;
            this.label17.Text = "ショートカット";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label22);
            this.panel1.Controls.Add(this.textBox5);
            this.panel1.Controls.Add(this.label21);
            this.panel1.Controls.Add(this.textBox4);
            this.panel1.Controls.Add(this.label20);
            this.panel1.Controls.Add(this.textBox3);
            this.panel1.Controls.Add(this.label19);
            this.panel1.Controls.Add(this.textBox2);
            this.panel1.Controls.Add(this.label18);
            this.panel1.Controls.Add(this.testPassedShortcutTextBox);
            this.panel1.Location = new System.Drawing.Point(17, 514);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(955, 209);
            this.panel1.TabIndex = 48;
            // 
            // label18
            // 
            this.label18.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(143, 13);
            this.label18.Margin = new System.Windows.Forms.Padding(0);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(111, 24);
            this.label18.TabIndex = 37;
            this.label18.Text = "テスト通過";
            this.label18.TextAlign = System.Drawing.ContentAlignment.BottomRight;
            // 
            // testPassedShortcutTextBox
            // 
            this.testPassedShortcutTextBox.Location = new System.Drawing.Point(257, 10);
            this.testPassedShortcutTextBox.Name = "testPassedShortcutTextBox";
            this.testPassedShortcutTextBox.Size = new System.Drawing.Size(482, 31);
            this.testPassedShortcutTextBox.TabIndex = 36;
            this.testPassedShortcutTextBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.testPassedShortcutTextBox_KeyDown);
            // 
            // label19
            // 
            this.label19.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(143, 50);
            this.label19.Margin = new System.Windows.Forms.Padding(0);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(111, 24);
            this.label19.TabIndex = 39;
            this.label19.Text = "テスト不可";
            this.label19.TextAlign = System.Drawing.ContentAlignment.BottomRight;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(257, 47);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(482, 31);
            this.textBox2.TabIndex = 38;
            // 
            // label20
            // 
            this.label20.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(14, 87);
            this.label20.Margin = new System.Windows.Forms.Padding(0);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(240, 24);
            this.label20.TabIndex = 41;
            this.label20.Text = "テスト通過(エビデンス有)";
            this.label20.TextAlign = System.Drawing.ContentAlignment.BottomRight;
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(257, 84);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(482, 31);
            this.textBox3.TabIndex = 40;
            // 
            // label21
            // 
            this.label21.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(14, 124);
            this.label21.Margin = new System.Windows.Forms.Padding(0);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(240, 24);
            this.label21.TabIndex = 43;
            this.label21.Text = "テスト不可(エビデンス有)";
            this.label21.TextAlign = System.Drawing.ContentAlignment.BottomRight;
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(257, 121);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(482, 31);
            this.textBox4.TabIndex = 42;
            // 
            // label22
            // 
            this.label22.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label22.AutoSize = true;
            this.label22.Location = new System.Drawing.Point(116, 161);
            this.label22.Margin = new System.Windows.Forms.Padding(0);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(133, 24);
            this.label22.TabIndex = 45;
            this.label22.Text = "テストスキップ";
            this.label22.TextAlign = System.Drawing.ContentAlignment.BottomRight;
            // 
            // textBox5
            // 
            this.textBox5.Location = new System.Drawing.Point(257, 158);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(482, 31);
            this.textBox5.TabIndex = 44;
            // 
            // MainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(13F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(991, 858);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.excelSettingsLabelLayoutPanel);
            this.Controls.Add(this.saveButton);
            this.Controls.Add(this.initializeButton);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainWindow";
            this.Text = "ExToolsForExcelTest";
            this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
            this.Load += new System.EventHandler(this.MainWindow_Load);
            this.contextMenuStrip.ResumeLayout(false);
            this.excelSettingsLabelLayoutPanel.ResumeLayout(false);
            this.excelSettingsLabelLayoutPanel.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.NotifyIcon notifyIcon;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip;
        private System.Windows.Forms.ToolStripMenuItem exitStripMenuItem;
        private System.Windows.Forms.Button saveButton;
        private System.Windows.Forms.Button initializeButton;
        private System.Windows.Forms.FlowLayoutPanel excelSettingsLabelLayoutPanel;
        private System.Windows.Forms.Label excelSettingsLabel;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label workbookKeyLabel;
        private System.Windows.Forms.TextBox workbookNameTextBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox testSheetNameTextBox;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox evidenceSheetTextBox;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox passedTextBox;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox ngTextBox;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox numberColumnNameTextBox;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox descriptionColumnNameTextBox;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox resultColumnNameTextBox;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox marginTopTextBox;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox imageMagnificationTextBox;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.TextBox testPassedShortcutTextBox;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.TextBox textBox5;
    }
}