using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TwinArch.SMRT_MVPLibrary.Interfaces;
using TwinArch.SMRT_MVPLibrary.Presenters;

namespace TwinArch.SMRT
{
    public partial class MainForm : Form, ISMRTMainView
    {
        private string fileName;
        private DataPresenter _presenter;

        public List<string> SheetNames
        {
            set { sheetNameCombo.DataSource = value; }
        }

        public Dictionary<string, string> ColumnNames
        {
            set
            {
                if ((value != null) && (value.Count > 0))
                {
                    columnsComboBox.DataSource = new BindingSource(value, null);
                    columnsComboBox.DisplayMember = "Value";
                    columnsComboBox.ValueMember = "Key";
                }
            }
        }

        public string AlertMessage
        {
            set { MessageBox.Show(value); }
        }


        public bool IsFileValid
        {
            set
            {
                if (!value)
                {
                    MessageBox.Show("The file selected is not a valid Excel for this tool. Double-check that the file exists, " +
                        "that it is not in use by any other program, that it is not read-only and you have permission to change it, " +
                        "and that it is an XLSX file (XLS files cannot be used).",
                    "SMRT - Invalid file",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                }
            }
        }

        public String UnhandledException
        {
            set
            {
                MessageBox.Show("There was an unexpected and unknown error. The message for the error is:\n"+ value,
                "SMRT - Unexpected error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
        }

        public MainForm()
        {
            InitializeComponent();
            sheetNameCombo.Enabled = false;
            columnsComboBox.Enabled = false;
            splitSourceButton.Enabled = false;
            getSheetsAndColumnsButton.Enabled = false;
            testTwitterButton.Enabled = false;
            _presenter = new DataPresenter(this, 3);
        }

        private void selectFileButton_Click(object sender, EventArgs e)
        {
            fileOpenDialog.FileName = excelFileNameTextBox.Text;

            fileOpenDialog.CheckFileExists = true;
            fileOpenDialog.Filter = "Excel files (*.xlsx, *.xls, *.xlsm)|*.xlsx;*.xls;*.xlsm|All files (*.*)|*.*";
            fileOpenDialog.Multiselect = false;

            if (fileOpenDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fileName = fileOpenDialog.FileName;
                excelFileNameTextBox.Text = fileName;
                getSheetsAndColumnsButton.Enabled = true;
            }
        }

        private void sheetNameCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            _presenter.DisplayColumnNames(excelFileNameTextBox.Text, sheetNameCombo.Text);
            columnsComboBox.Enabled = true;
            testTwitterButton.Enabled = true;
            this.Cursor = Cursors.Default;
        }

        private void splitSourceButton_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            SMRT_MVPLibrary.ReturnCode rc = _presenter.ParseURLs(fileName, sheetNameCombo.Text, columnsComboBox.Text, false, firstRowIsAColumnHeaderCheckBox.Checked);
            if (rc == SMRT_MVPLibrary.ReturnCode.ColumnsAlreadyExist)
            {
                this.Cursor = Cursors.Default;
                DialogResult result = MessageBox.Show("The columns that will contain the Domain Name, Poster ID, and Mention ID already exist in this Excel file. Do you want to overwrite the data that already exists in those columns?",
                    "SMRT - Columns already exist",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    this.Cursor = Cursors.WaitCursor;
                    rc = _presenter.ParseURLs(fileName, sheetNameCombo.Text, columnsComboBox.Text, true, firstRowIsAColumnHeaderCheckBox.Checked);
                    this.Cursor = Cursors.Default;
                }
            }
            if (rc == SMRT_MVPLibrary.ReturnCode.NotURLColumn)
            {
                this.Cursor = Cursors.Default;
                MessageBox.Show("The column you selected does not appear to contain URLs. Please select another column.",
                    "SMRT - Column does not contain URLs.",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            this.Cursor = Cursors.Default;
            if (rc == SMRT_MVPLibrary.ReturnCode.Success)
                MessageBox.Show("Done!", "SMRT - Split Source Into Parts", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void getSheetsAndColumnsButton_Click(object sender, EventArgs e)
        {
            _presenter.EmptySheetNames();
            _presenter.EmptyColumnNames();
            this.Cursor = Cursors.WaitCursor;
            _presenter.DisplaySheetNames(fileName);
            this.Cursor = Cursors.Default;
            sheetNameCombo.Enabled = true;
        }

        private void columnsComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            splitSourceButton.Enabled = true;
        }

        private void testTwitterButton_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            SMRT_MVPLibrary.ReturnCode rc = _presenter.AddTwitterInfo(fileName, sheetNameCombo.Text, false, firstRowIsAColumnHeaderCheckBox.Checked);
            if (rc == SMRT_MVPLibrary.ReturnCode.ColumnsAlreadyExist)
            {
                this.Cursor = Cursors.Default;
                DialogResult result = MessageBox.Show("The columns that will contain the Twitter info already exist in this Excel file. Do you want to overwrite the data that already exists in those columns?",
                    "SMRT - Columns already exist",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    this.Cursor = Cursors.WaitCursor;
                    rc = _presenter.AddTwitterInfo(fileName, sheetNameCombo.Text, true, firstRowIsAColumnHeaderCheckBox.Checked);
                    this.Cursor = Cursors.Default;
                }
            }
            if (rc == SMRT_MVPLibrary.ReturnCode.ColumnsMissing)
            {
                this.Cursor = Cursors.Default;
                MessageBox.Show("The sheet you selected does not appear to contain a column called PosterID."+
                    " This column is created when you run the \"Split Source Into Parts\" function, so please run that first.",
                    "SMRT - Sheet does not contain PosterID.",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            this.Cursor = Cursors.Default;
            if (rc == SMRT_MVPLibrary.ReturnCode.Success)
                MessageBox.Show("Done!", "SMRT - Get Twitter Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


    }
}
