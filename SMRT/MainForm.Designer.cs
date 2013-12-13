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
            this.numOfTopPostersTextBox = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.tabControl1.SuspendLayout();
            this.tabSplit.SuspendLayout();
            this.tabTwitter.SuspendLayout();
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
            this.tabControl1.Location = new System.Drawing.Point(16, 97);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(445, 126);
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
            this.tabSplit.Size = new System.Drawing.Size(404, 100);
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
            this.tabTwitter.Size = new System.Drawing.Size(437, 100);
            this.tabTwitter.TabIndex = 1;
            this.tabTwitter.Text = "Get Twitter Info";
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
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(7, 9);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(167, 13);
            this.label4.TabIndex = 15;
            this.label4.Text = "Number of top posters to retrieve: ";
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
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(481, 262);
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
    }
}

