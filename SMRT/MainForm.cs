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

        public MainForm()
        {
            InitializeComponent();
            sheetNameCombo.Enabled = false;
            columnsComboBox.Enabled = false;
            testOneButton.Enabled = false;
            testTwoButton.Enabled = false;
            testThreeButton.Enabled = false;
            testURLParseButton.Enabled = false;
            getSheetsAndColumnsButton.Enabled = false;
        }

        private void selectFileButton_Click(object sender, EventArgs e)
        {
            fileOpenDialog.FileName = excelFileNameTextBox.Text;

            fileOpenDialog.CheckFileExists = true;
            fileOpenDialog.Filter = "Excel files (*.xlsx, *.xls)|*.xlsx;*.xls|All files (*.*)|*.*";
            fileOpenDialog.Multiselect = false;

            if (fileOpenDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fileName = fileOpenDialog.FileName;
                excelFileNameTextBox.Text = fileName;

                testOneButton.Enabled = true;
                testTwoButton.Enabled = true;
                testThreeButton.Enabled = true;
            }
        }

        private void sheetNameCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            _presenter.DisplayColumnNames(excelFileNameTextBox.Text, sheetNameCombo.Text);
            columnsComboBox.Enabled = true;
        }

        private void testOneButton_Click(object sender, EventArgs e)
        {
            if (_presenter != null) _presenter.Dispose();
            _presenter = new DataPresenter(this, 0);
            getSheetsAndColumnsButton.Enabled = true;
        }

        private void testTwoButton_Click(object sender, EventArgs e)
        {
            if (_presenter != null) _presenter.Dispose();
            _presenter = new DataPresenter(this, 1);
            getSheetsAndColumnsButton.Enabled = true;
        }

        private void testThreeButton_Click(object sender, EventArgs e)
        {
            if (_presenter != null) _presenter.Dispose();
            _presenter = new DataPresenter(this, 3);
            getSheetsAndColumnsButton.Enabled = true;
        }

        private void testURLParseButton_Click(object sender, EventArgs e)
        {
            SMRT_MVPLibrary.ReturnCode rc = _presenter.ParseURLs(fileName, sheetNameCombo.Text, columnsComboBox.SelectedValue.ToString(), false);
            if (rc == SMRT_MVPLibrary.ReturnCode.ColumnsAlreadyExist)
            {
                DialogResult result = MessageBox.Show("The columns into which the parts of the URL are to be placed already exist in this Excel file. Do you want to overwrite the data that already exists in those columns?",
                    "Columns already exist",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                    rc = _presenter.ParseURLs(fileName, sheetNameCombo.Text, columnsComboBox.SelectedValue.ToString(), true);
            }
            if (rc == SMRT_MVPLibrary.ReturnCode.NotURLColumn)
                MessageBox.Show("The column you selected does not appear to contain URLs. Please select another column.",
                    "Column does not contain URLs.",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
        }

        private void getSheetsAndColumnsButton_Click(object sender, EventArgs e)
        {
            _presenter.EmptySheetNames();
            _presenter.EmptyColumnNames();
            _presenter.DisplaySheetNames(fileName);
            sheetNameCombo.Enabled = true;
        }

        private void columnsComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            testURLParseButton.Enabled = true;
        }

    }
}
