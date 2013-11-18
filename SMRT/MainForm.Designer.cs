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
            this.label1 = new System.Windows.Forms.Label();
            this.excelFileNameTextBox = new System.Windows.Forms.TextBox();
            this.selectFileButton = new System.Windows.Forms.Button();
            this.fileOpenDialog = new System.Windows.Forms.OpenFileDialog();
            this.sheetNameCombo = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.columnsComboBox = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.testOneButton = new System.Windows.Forms.Button();
            this.testTwoButton = new System.Windows.Forms.Button();
            this.testThreeButton = new System.Windows.Forms.Button();
            this.testURLParseButton = new System.Windows.Forms.Button();
            this.getSheetsAndColumnsButton = new System.Windows.Forms.Button();
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
            this.excelFileNameTextBox.Location = new System.Drawing.Point(148, 13);
            this.excelFileNameTextBox.Name = "excelFileNameTextBox";
            this.excelFileNameTextBox.Size = new System.Drawing.Size(264, 20);
            this.excelFileNameTextBox.TabIndex = 1;
            // 
            // selectFileButton
            // 
            this.selectFileButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.selectFileButton.Location = new System.Drawing.Point(419, 9);
            this.selectFileButton.Name = "selectFileButton";
            this.selectFileButton.Size = new System.Drawing.Size(75, 23);
            this.selectFileButton.TabIndex = 2;
            this.selectFileButton.Text = "Select file...";
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
            this.sheetNameCombo.Location = new System.Drawing.Point(148, 97);
            this.sheetNameCombo.Name = "sheetNameCombo";
            this.sheetNameCombo.Size = new System.Drawing.Size(264, 21);
            this.sheetNameCombo.TabIndex = 3;
            this.sheetNameCombo.SelectedIndexChanged += new System.EventHandler(this.sheetNameCombo_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 100);
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
            this.columnsComboBox.Location = new System.Drawing.Point(148, 125);
            this.columnsComboBox.Name = "columnsComboBox";
            this.columnsComboBox.Size = new System.Drawing.Size(264, 21);
            this.columnsComboBox.TabIndex = 5;
            this.columnsComboBox.SelectedIndexChanged += new System.EventHandler(this.columnsComboBox_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 128);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(129, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Select column with URLs:";
            // 
            // testOneButton
            // 
            this.testOneButton.Location = new System.Drawing.Point(148, 39);
            this.testOneButton.Name = "testOneButton";
            this.testOneButton.Size = new System.Drawing.Size(122, 23);
            this.testOneButton.TabIndex = 7;
            this.testOneButton.Text = "Test Excel Automation";
            this.testOneButton.UseVisualStyleBackColor = true;
            this.testOneButton.Click += new System.EventHandler(this.testOneButton_Click);
            // 
            // testTwoButton
            // 
            this.testTwoButton.Location = new System.Drawing.Point(276, 39);
            this.testTwoButton.Name = "testTwoButton";
            this.testTwoButton.Size = new System.Drawing.Size(97, 23);
            this.testTwoButton.TabIndex = 8;
            this.testTwoButton.Text = "Test OLEDB Jet";
            this.testTwoButton.UseVisualStyleBackColor = true;
            this.testTwoButton.Click += new System.EventHandler(this.testTwoButton_Click);
            // 
            // testThreeButton
            // 
            this.testThreeButton.Location = new System.Drawing.Point(379, 39);
            this.testThreeButton.Name = "testThreeButton";
            this.testThreeButton.Size = new System.Drawing.Size(114, 23);
            this.testThreeButton.TabIndex = 9;
            this.testThreeButton.Text = "Test OLEDB ACE";
            this.testThreeButton.UseVisualStyleBackColor = true;
            this.testThreeButton.Click += new System.EventHandler(this.testThreeButton_Click);
            // 
            // testURLParseButton
            // 
            this.testURLParseButton.Location = new System.Drawing.Point(148, 152);
            this.testURLParseButton.Name = "testURLParseButton";
            this.testURLParseButton.Size = new System.Drawing.Size(91, 23);
            this.testURLParseButton.TabIndex = 10;
            this.testURLParseButton.Text = "Test URL Parse";
            this.testURLParseButton.UseVisualStyleBackColor = true;
            this.testURLParseButton.Click += new System.EventHandler(this.testURLParseButton_Click);
            // 
            // getSheetsAndColumnsButton
            // 
            this.getSheetsAndColumnsButton.Location = new System.Drawing.Point(148, 68);
            this.getSheetsAndColumnsButton.Name = "getSheetsAndColumnsButton";
            this.getSheetsAndColumnsButton.Size = new System.Drawing.Size(122, 23);
            this.getSheetsAndColumnsButton.TabIndex = 11;
            this.getSheetsAndColumnsButton.Text = "Get Sheets/Columns";
            this.getSheetsAndColumnsButton.UseVisualStyleBackColor = true;
            this.getSheetsAndColumnsButton.Click += new System.EventHandler(this.getSheetsAndColumnsButton_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(505, 262);
            this.Controls.Add(this.getSheetsAndColumnsButton);
            this.Controls.Add(this.testURLParseButton);
            this.Controls.Add(this.testThreeButton);
            this.Controls.Add(this.testTwoButton);
            this.Controls.Add(this.testOneButton);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.columnsComboBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.sheetNameCombo);
            this.Controls.Add(this.selectFileButton);
            this.Controls.Add(this.excelFileNameTextBox);
            this.Controls.Add(this.label1);
            this.Name = "MainForm";
            this.Text = "SMRT - Social Media Research Tool";
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
        private System.Windows.Forms.Button testOneButton;
        private System.Windows.Forms.Button testTwoButton;
        private System.Windows.Forms.Button testThreeButton;
        private System.Windows.Forms.Button testURLParseButton;
        private System.Windows.Forms.Button getSheetsAndColumnsButton;
    }
}

