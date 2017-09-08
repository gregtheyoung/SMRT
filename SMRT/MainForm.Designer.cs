namespace TwinArch.SMRT
{
    partial class MainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.label1 = new System.Windows.Forms.Label();
            this.excelFileNameTextBox = new System.Windows.Forms.TextBox();
            this.selectFileButton = new System.Windows.Forms.Button();
            this.fileOpenDialog = new System.Windows.Forms.OpenFileDialog();
            this.sheetNameCombo = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.columnsComboBox = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.splitSourceButton = new System.Windows.Forms.Button();
            this.getSheetsAndColumnsButton = new System.Windows.Forms.Button();
            this.firstRowIsAColumnHeaderCheckBox = new System.Windows.Forms.CheckBox();
            this.testTwitterButton = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabSplit = new System.Windows.Forms.TabPage();
            this.tabTwitter = new System.Windows.Forms.TabPage();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.numOfTopPostersTextBox = new System.Windows.Forms.TextBox();
            this.tabAutocode = new System.Windows.Forms.TabPage();
            this.ignoreSecondColumnCheckbox = new System.Windows.Forms.CheckBox();
            this.columnsForHostStringTextComboBox = new System.Windows.Forms.ComboBox();
            this.label13 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.buttonAutocode = new System.Windows.Forms.Button();
            this.selectCodeFamilyFileButton = new System.Windows.Forms.Button();
            this.codeFamilyFileNameTextBox = new System.Windows.Forms.TextBox();
            this.columnsForMentionTextComboBox = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.tabRandomSelect = new System.Windows.Forms.TabPage();
            this.ceilingNumeric = new System.Windows.Forms.NumericUpDown();
            this.floorNumeric = new System.Windows.Forms.NumericUpDown();
            this.percentageNumeric = new System.Windows.Forms.NumericUpDown();
            this.label12 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.columnsForRandomSelectComboxBox = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.randomSelectButton = new System.Windows.Forms.Button();
            this.columnsForAutocodeCountsComboxBox = new System.Windows.Forms.ComboBox();
            this.oneGroupRadioButton = new System.Windows.Forms.RadioButton();
            this.multipleGroupRadioButton = new System.Windows.Forms.RadioButton();
            this.label14 = new System.Windows.Forms.Label();
            this.numberOfGroupsNumeric = new System.Windows.Forms.NumericUpDown();
            this.tabControl1.SuspendLayout();
            this.tabSplit.SuspendLayout();
            this.tabTwitter.SuspendLayout();
            this.tabAutocode.SuspendLayout();
            this.tabRandomSelect.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ceilingNumeric)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.floorNumeric)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.percentageNumeric)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numberOfGroupsNumeric)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(88, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Select Excel file: ";
            // 
            // excelFileNameTextBox
            // 
            this.excelFileNameTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.excelFileNameTextBox.Location = new System.Drawing.Point(107, 13);
            this.excelFileNameTextBox.Name = "excelFileNameTextBox";
            this.excelFileNameTextBox.Size = new System.Drawing.Size(313, 20);
            this.excelFileNameTextBox.TabIndex = 1;
            // 
            // selectFileButton
            // 
            this.selectFileButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.selectFileButton.Location = new System.Drawing.Point(426, 13);
            this.selectFileButton.Name = "selectFileButton";
            this.selectFileButton.Size = new System.Drawing.Size(28, 23);
            this.selectFileButton.TabIndex = 2;
            this.selectFileButton.Text = "...";
            this.selectFileButton.UseVisualStyleBackColor = true;
            this.selectFileButton.Click += new System.EventHandler(this.selectFileButton_Click);
            // 
            // fileOpenDialog
            // 
            this.fileOpenDialog.FileName = "openFileDialog1";
            // 
            // sheetNameCombo
            // 
            this.sheetNameCombo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.sheetNameCombo.FormattingEnabled = true;
            this.sheetNameCombo.Location = new System.Drawing.Point(107, 70);
            this.sheetNameCombo.Name = "sheetNameCombo";
            this.sheetNameCombo.Size = new System.Drawing.Size(313, 21);
            this.sheetNameCombo.TabIndex = 3;
            this.sheetNameCombo.SelectedIndexChanged += new System.EventHandler(this.sheetNameCombo_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 73);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(69, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Select sheet:";
            // 
            // columnsComboBox
            // 
            this.columnsComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.columnsComboBox.FormattingEnabled = true;
            this.columnsComboBox.Location = new System.Drawing.Point(176, 12);
            this.columnsComboBox.Name = "columnsComboBox";
            this.columnsComboBox.Size = new System.Drawing.Size(157, 21);
            this.columnsComboBox.TabIndex = 5;
            this.columnsComboBox.SelectedIndexChanged += new System.EventHandler(this.columnsComboBox_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 15);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(164, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Select column with source URLs:";
            // 
            // splitSourceButton
            // 
            this.splitSourceButton.Location = new System.Drawing.Point(9, 62);
            this.splitSourceButton.Name = "splitSourceButton";
            this.splitSourceButton.Size = new System.Drawing.Size(122, 23);
            this.splitSourceButton.TabIndex = 10;
            this.splitSourceButton.Text = "Split Source Into Parts";
            this.splitSourceButton.UseVisualStyleBackColor = true;
            this.splitSourceButton.Click += new System.EventHandler(this.splitSourceButton_Click);
            // 
            // getSheetsAndColumnsButton
            // 
            this.getSheetsAndColumnsButton.Location = new System.Drawing.Point(107, 39);
            this.getSheetsAndColumnsButton.Name = "getSheetsAndColumnsButton";
            this.getSheetsAndColumnsButton.Size = new System.Drawing.Size(122, 23);
            this.getSheetsAndColumnsButton.TabIndex = 11;
            this.getSheetsAndColumnsButton.Text = "Get Sheets/Columns";
            this.getSheetsAndColumnsButton.UseVisualStyleBackColor = true;
            this.getSheetsAndColumnsButton.Click += new System.EventHandler(this.getSheetsAndColumnsButton_Click);
            // 
            // firstRowIsAColumnHeaderCheckBox
            // 
            this.firstRowIsAColumnHeaderCheckBox.AutoSize = true;
            this.firstRowIsAColumnHeaderCheckBox.Checked = true;
            this.firstRowIsAColumnHeaderCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.firstRowIsAColumnHeaderCheckBox.Location = new System.Drawing.Point(9, 39);
            this.firstRowIsAColumnHeaderCheckBox.Name = "firstRowIsAColumnHeaderCheckBox";
            this.firstRowIsAColumnHeaderCheckBox.Size = new System.Drawing.Size(241, 17);
            this.firstRowIsAColumnHeaderCheckBox.TabIndex = 12;
            this.firstRowIsAColumnHeaderCheckBox.Text = "The first row in this sheet has column headers";
            this.firstRowIsAColumnHeaderCheckBox.UseVisualStyleBackColor = true;
            // 
            // testTwitterButton
            // 
            this.testTwitterButton.Location = new System.Drawing.Point(6, 35);
            this.testTwitterButton.Name = "testTwitterButton";
            this.testTwitterButton.Size = new System.Drawing.Size(122, 23);
            this.testTwitterButton.TabIndex = 13;
            this.testTwitterButton.Text = "Get Twitter Info";
            this.testTwitterButton.UseVisualStyleBackColor = true;
            this.testTwitterButton.Click += new System.EventHandler(this.testTwitterButton_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabSplit);
            this.tabControl1.Controls.Add(this.tabTwitter);
            this.tabControl1.Controls.Add(this.tabAutocode);
            this.tabControl1.Controls.Add(this.tabRandomSelect);
            this.tabControl1.Location = new System.Drawing.Point(16, 97);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(445, 207);
            this.tabControl1.TabIndex = 14;
            // 
            // tabSplit
            // 
            this.tabSplit.BackColor = System.Drawing.Color.Transparent;
            this.tabSplit.Controls.Add(this.firstRowIsAColumnHeaderCheckBox);
            this.tabSplit.Controls.Add(this.columnsComboBox);
            this.tabSplit.Controls.Add(this.label3);
            this.tabSplit.Controls.Add(this.splitSourceButton);
            this.tabSplit.Location = new System.Drawing.Point(4, 22);
            this.tabSplit.Name = "tabSplit";
            this.tabSplit.Padding = new System.Windows.Forms.Padding(3);
            this.tabSplit.Size = new System.Drawing.Size(437, 181);
            this.tabSplit.TabIndex = 0;
            this.tabSplit.Text = "Split Source Into Parts";
            // 
            // tabTwitter
            // 
            this.tabTwitter.BackColor = System.Drawing.SystemColors.Control;
            this.tabTwitter.Controls.Add(this.label5);
            this.tabTwitter.Controls.Add(this.label4);
            this.tabTwitter.Controls.Add(this.numOfTopPostersTextBox);
            this.tabTwitter.Controls.Add(this.testTwitterButton);
            this.tabTwitter.Location = new System.Drawing.Point(4, 22);
            this.tabTwitter.Name = "tabTwitter";
            this.tabTwitter.Padding = new System.Windows.Forms.Padding(3);
            this.tabTwitter.Size = new System.Drawing.Size(437, 181);
            this.tabTwitter.TabIndex = 1;
            this.tabTwitter.Text = "Get Twitter Info";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(236, 9);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(195, 13);
            this.label5.TabIndex = 16;
            this.label5.Text = "(do not exceed 180 due to Twitter limits)";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(7, 9);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(167, 13);
            this.label4.TabIndex = 15;
            this.label4.Text = "Number of top posters to retrieve: ";
            // 
            // numOfTopPostersTextBox
            // 
            this.numOfTopPostersTextBox.Location = new System.Drawing.Point(177, 6);
            this.numOfTopPostersTextBox.Name = "numOfTopPostersTextBox";
            this.numOfTopPostersTextBox.Size = new System.Drawing.Size(53, 20);
            this.numOfTopPostersTextBox.TabIndex = 14;
            this.numOfTopPostersTextBox.TextChanged += new System.EventHandler(this.numOfTopPostersTextBox_TextChanged);
            this.numOfTopPostersTextBox.Leave += new System.EventHandler(this.numOfTopPostersTextBox_Leave);
            // 
            // tabAutocode
            // 
            this.tabAutocode.BackColor = System.Drawing.SystemColors.Control;
            this.tabAutocode.Controls.Add(this.ignoreSecondColumnCheckbox);
            this.tabAutocode.Controls.Add(this.columnsForHostStringTextComboBox);
            this.tabAutocode.Controls.Add(this.label13);
            this.tabAutocode.Controls.Add(this.label8);
            this.tabAutocode.Controls.Add(this.buttonAutocode);
            this.tabAutocode.Controls.Add(this.selectCodeFamilyFileButton);
            this.tabAutocode.Controls.Add(this.codeFamilyFileNameTextBox);
            this.tabAutocode.Controls.Add(this.columnsForMentionTextComboBox);
            this.tabAutocode.Controls.Add(this.label6);
            this.tabAutocode.Location = new System.Drawing.Point(4, 22);
            this.tabAutocode.Name = "tabAutocode";
            this.tabAutocode.Padding = new System.Windows.Forms.Padding(3);
            this.tabAutocode.Size = new System.Drawing.Size(437, 154);
            this.tabAutocode.TabIndex = 2;
            this.tabAutocode.Text = "Autocode";
            // 
            // ignoreSecondColumnCheckbox
            // 
            this.ignoreSecondColumnCheckbox.AutoSize = true;
            this.ignoreSecondColumnCheckbox.Location = new System.Drawing.Point(378, 38);
            this.ignoreSecondColumnCheckbox.Name = "ignoreSecondColumnCheckbox";
            this.ignoreSecondColumnCheckbox.Size = new System.Drawing.Size(56, 17);
            this.ignoreSecondColumnCheckbox.TabIndex = 4;
            this.ignoreSecondColumnCheckbox.Text = "Ignore";
            this.ignoreSecondColumnCheckbox.UseVisualStyleBackColor = true;
            this.ignoreSecondColumnCheckbox.CheckedChanged += new System.EventHandler(this.ignoreSecondColumnCheckbox_CheckedChanged);
            // 
            // columnsForHostStringTextComboBox
            // 
            this.columnsForHostStringTextComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.columnsForHostStringTextComboBox.FormattingEnabled = true;
            this.columnsForHostStringTextComboBox.Location = new System.Drawing.Point(214, 35);
            this.columnsForHostStringTextComboBox.Name = "columnsForHostStringTextComboBox";
            this.columnsForHostStringTextComboBox.Size = new System.Drawing.Size(157, 21);
            this.columnsForHostStringTextComboBox.TabIndex = 3;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(6, 38);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(193, 13);
            this.label13.TabIndex = 2;
            this.label13.Text = "Select second column (host/link string):";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(6, 70);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(127, 13);
            this.label8.TabIndex = 5;
            this.label8.Text = "Select Code Families file: ";
            this.label8.Click += new System.EventHandler(this.label8_Click);
            // 
            // buttonAutocode
            // 
            this.buttonAutocode.Location = new System.Drawing.Point(9, 93);
            this.buttonAutocode.Name = "buttonAutocode";
            this.buttonAutocode.Size = new System.Drawing.Size(122, 23);
            this.buttonAutocode.TabIndex = 8;
            this.buttonAutocode.Text = "Autocode";
            this.buttonAutocode.UseVisualStyleBackColor = true;
            this.buttonAutocode.Click += new System.EventHandler(this.buttonAutocode_Click);
            // 
            // selectCodeFamilyFileButton
            // 
            this.selectCodeFamilyFileButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.selectCodeFamilyFileButton.Location = new System.Drawing.Point(406, 67);
            this.selectCodeFamilyFileButton.Name = "selectCodeFamilyFileButton";
            this.selectCodeFamilyFileButton.Size = new System.Drawing.Size(28, 23);
            this.selectCodeFamilyFileButton.TabIndex = 7;
            this.selectCodeFamilyFileButton.Text = "...";
            this.selectCodeFamilyFileButton.UseVisualStyleBackColor = true;
            this.selectCodeFamilyFileButton.Click += new System.EventHandler(this.selectCodeFamilyFileButton_Click);
            // 
            // codeFamilyFileNameTextBox
            // 
            this.codeFamilyFileNameTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.codeFamilyFileNameTextBox.Location = new System.Drawing.Point(130, 67);
            this.codeFamilyFileNameTextBox.Name = "codeFamilyFileNameTextBox";
            this.codeFamilyFileNameTextBox.Size = new System.Drawing.Size(270, 20);
            this.codeFamilyFileNameTextBox.TabIndex = 6;
            // 
            // columnsForMentionTextComboBox
            // 
            this.columnsForMentionTextComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.columnsForMentionTextComboBox.FormattingEnabled = true;
            this.columnsForMentionTextComboBox.Location = new System.Drawing.Point(214, 10);
            this.columnsForMentionTextComboBox.Name = "columnsForMentionTextComboBox";
            this.columnsForMentionTextComboBox.Size = new System.Drawing.Size(157, 21);
            this.columnsForMentionTextComboBox.TabIndex = 1;
            this.columnsForMentionTextComboBox.SelectedIndexChanged += new System.EventHandler(this.columnsForMentionTextComboBox_SelectedIndexChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(6, 13);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(162, 13);
            this.label6.TabIndex = 0;
            this.label6.Text = "Select first column (mention text):";
            // 
            // tabRandomSelect
            // 
            this.tabRandomSelect.BackColor = System.Drawing.SystemColors.Control;
            this.tabRandomSelect.Controls.Add(this.numberOfGroupsNumeric);
            this.tabRandomSelect.Controls.Add(this.label14);
            this.tabRandomSelect.Controls.Add(this.multipleGroupRadioButton);
            this.tabRandomSelect.Controls.Add(this.oneGroupRadioButton);
            this.tabRandomSelect.Controls.Add(this.ceilingNumeric);
            this.tabRandomSelect.Controls.Add(this.floorNumeric);
            this.tabRandomSelect.Controls.Add(this.percentageNumeric);
            this.tabRandomSelect.Controls.Add(this.label12);
            this.tabRandomSelect.Controls.Add(this.label11);
            this.tabRandomSelect.Controls.Add(this.label10);
            this.tabRandomSelect.Controls.Add(this.label9);
            this.tabRandomSelect.Controls.Add(this.columnsForRandomSelectComboxBox);
            this.tabRandomSelect.Controls.Add(this.label7);
            this.tabRandomSelect.Controls.Add(this.randomSelectButton);
            this.tabRandomSelect.Controls.Add(this.columnsForAutocodeCountsComboxBox);
            this.tabRandomSelect.Location = new System.Drawing.Point(4, 22);
            this.tabRandomSelect.Name = "tabRandomSelect";
            this.tabRandomSelect.Size = new System.Drawing.Size(437, 181);
            this.tabRandomSelect.TabIndex = 3;
            this.tabRandomSelect.Text = "Random Select";
            // 
            // ceilingNumeric
            // 
            this.ceilingNumeric.Location = new System.Drawing.Point(90, 119);
            this.ceilingNumeric.Maximum = new decimal(new int[] {
            100000,
            0,
            0,
            0});
            this.ceilingNumeric.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.ceilingNumeric.Name = "ceilingNumeric";
            this.ceilingNumeric.Size = new System.Drawing.Size(41, 20);
            this.ceilingNumeric.TabIndex = 6;
            this.ceilingNumeric.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // floorNumeric
            // 
            this.floorNumeric.Location = new System.Drawing.Point(90, 97);
            this.floorNumeric.Maximum = new decimal(new int[] {
            100000,
            0,
            0,
            0});
            this.floorNumeric.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.floorNumeric.Name = "floorNumeric";
            this.floorNumeric.Size = new System.Drawing.Size(41, 20);
            this.floorNumeric.TabIndex = 5;
            this.floorNumeric.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // percentageNumeric
            // 
            this.percentageNumeric.Location = new System.Drawing.Point(90, 75);
            this.percentageNumeric.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.percentageNumeric.Name = "percentageNumeric";
            this.percentageNumeric.Size = new System.Drawing.Size(41, 20);
            this.percentageNumeric.TabIndex = 4;
            this.percentageNumeric.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(3, 121);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(69, 13);
            this.label12.TabIndex = 6;
            this.label12.Text = "Celing count:";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(3, 99);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(63, 13);
            this.label11.TabIndex = 5;
            this.label11.Text = "Floor count:";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(3, 77);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(81, 13);
            this.label10.TabIndex = 4;
            this.label10.Text = "Target Percent:";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(3, 31);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(175, 13);
            this.label9.TabIndex = 23;
            this.label9.Text = "Select column for random selection:";
            // 
            // columnsForRandomSelectComboxBox
            // 
            this.columnsForRandomSelectComboxBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.columnsForRandomSelectComboxBox.FormattingEnabled = true;
            this.columnsForRandomSelectComboxBox.Location = new System.Drawing.Point(184, 28);
            this.columnsForRandomSelectComboxBox.Name = "columnsForRandomSelectComboxBox";
            this.columnsForRandomSelectComboxBox.Size = new System.Drawing.Size(157, 21);
            this.columnsForRandomSelectComboxBox.TabIndex = 1;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(3, 10);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(182, 13);
            this.label7.TabIndex = 21;
            this.label7.Text = "Select column with autocode counts:";
            // 
            // randomSelectButton
            // 
            this.randomSelectButton.Location = new System.Drawing.Point(6, 148);
            this.randomSelectButton.Name = "randomSelectButton";
            this.randomSelectButton.Size = new System.Drawing.Size(122, 23);
            this.randomSelectButton.TabIndex = 8;
            this.randomSelectButton.Text = "Random Select";
            this.randomSelectButton.UseVisualStyleBackColor = true;
            this.randomSelectButton.Click += new System.EventHandler(this.randomSelectButton_Click);
            // 
            // columnsForAutocodeCountsComboxBox
            // 
            this.columnsForAutocodeCountsComboxBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.columnsForAutocodeCountsComboxBox.FormattingEnabled = true;
            this.columnsForAutocodeCountsComboxBox.Location = new System.Drawing.Point(184, 7);
            this.columnsForAutocodeCountsComboxBox.Name = "columnsForAutocodeCountsComboxBox";
            this.columnsForAutocodeCountsComboxBox.Size = new System.Drawing.Size(157, 21);
            this.columnsForAutocodeCountsComboxBox.TabIndex = 0;
            // 
            // oneGroupRadioButton
            // 
            this.oneGroupRadioButton.AutoSize = true;
            this.oneGroupRadioButton.Checked = true;
            this.oneGroupRadioButton.Location = new System.Drawing.Point(6, 52);
            this.oneGroupRadioButton.Name = "oneGroupRadioButton";
            this.oneGroupRadioButton.Size = new System.Drawing.Size(131, 17);
            this.oneGroupRadioButton.TabIndex = 2;
            this.oneGroupRadioButton.TabStop = true;
            this.oneGroupRadioButton.Text = "One Group by Percent";
            this.oneGroupRadioButton.UseVisualStyleBackColor = true;
            // 
            // multipleGroupRadioButton
            // 
            this.multipleGroupRadioButton.AutoSize = true;
            this.multipleGroupRadioButton.Location = new System.Drawing.Point(184, 52);
            this.multipleGroupRadioButton.Name = "multipleGroupRadioButton";
            this.multipleGroupRadioButton.Size = new System.Drawing.Size(98, 17);
            this.multipleGroupRadioButton.TabIndex = 3;
            this.multipleGroupRadioButton.Text = "Multiple Groups";
            this.multipleGroupRadioButton.UseVisualStyleBackColor = true;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(181, 77);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(96, 13);
            this.label14.TabIndex = 34;
            this.label14.Text = "Number of Groups:";
            // 
            // numberOfGroupsNumeric
            // 
            this.numberOfGroupsNumeric.Location = new System.Drawing.Point(284, 74);
            this.numberOfGroupsNumeric.Name = "numberOfGroupsNumeric";
            this.numberOfGroupsNumeric.Size = new System.Drawing.Size(57, 20);
            this.numberOfGroupsNumeric.TabIndex = 7;
            this.numberOfGroupsNumeric.Value = new decimal(new int[] {
            10,
            0,
            0,
            0});
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(481, 316);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.getSheetsAndColumnsButton);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.sheetNameCombo);
            this.Controls.Add(this.selectFileButton);
            this.Controls.Add(this.excelFileNameTextBox);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainForm";
            this.Text = "SMRT - Social Media Research Tool";
            this.tabControl1.ResumeLayout(false);
            this.tabSplit.ResumeLayout(false);
            this.tabSplit.PerformLayout();
            this.tabTwitter.ResumeLayout(false);
            this.tabTwitter.PerformLayout();
            this.tabAutocode.ResumeLayout(false);
            this.tabAutocode.PerformLayout();
            this.tabRandomSelect.ResumeLayout(false);
            this.tabRandomSelect.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ceilingNumeric)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.floorNumeric)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.percentageNumeric)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numberOfGroupsNumeric)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox excelFileNameTextBox;
        private System.Windows.Forms.Button selectFileButton;
        private System.Windows.Forms.OpenFileDialog fileOpenDialog;
        private System.Windows.Forms.ComboBox sheetNameCombo;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox columnsComboBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button splitSourceButton;
        private System.Windows.Forms.Button getSheetsAndColumnsButton;
        private System.Windows.Forms.CheckBox firstRowIsAColumnHeaderCheckBox;
        private System.Windows.Forms.Button testTwitterButton;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabSplit;
        private System.Windows.Forms.TabPage tabTwitter;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox numOfTopPostersTextBox;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TabPage tabAutocode;
        private System.Windows.Forms.ComboBox columnsForMentionTextComboBox;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button buttonAutocode;
        private System.Windows.Forms.Button selectCodeFamilyFileButton;
        private System.Windows.Forms.TextBox codeFamilyFileNameTextBox;
        private System.Windows.Forms.TabPage tabRandomSelect;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox columnsForRandomSelectComboxBox;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button randomSelectButton;
        private System.Windows.Forms.ComboBox columnsForAutocodeCountsComboxBox;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.NumericUpDown ceilingNumeric;
        private System.Windows.Forms.NumericUpDown floorNumeric;
        private System.Windows.Forms.NumericUpDown percentageNumeric;
        private System.Windows.Forms.ComboBox columnsForHostStringTextComboBox;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.CheckBox ignoreSecondColumnCheckbox;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.RadioButton multipleGroupRadioButton;
        private System.Windows.Forms.RadioButton oneGroupRadioButton;
        private System.Windows.Forms.NumericUpDown numberOfGroupsNumeric;
    }
}

