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

        public List<string> ColumnNames
        {
            set { columnsComboBox.DataSource = value; }
        }

        public MainForm()
        {
            InitializeComponent();
            sheetNameCombo.Enabled = false;
            columnsComboBox.Enabled = false;
            testOneButton.Enabled = false;
            testTwoButton.Enabled = false;
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
            }
        }

        private void sheetNameCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            _presenter.DisplayColumnNames(excelFileNameTextBox.Text, sheetNameCombo.Text);
            columnsComboBox.Enabled = true;
        }

        private void testOneButton_Click(object sender, EventArgs e)
        {
            _presenter = new DataPresenter(this, true);
            _presenter.EmptySheetNames();
            _presenter.EmptyColumnNames();
            _presenter.DisplaySheetNames(fileName);
            sheetNameCombo.Enabled = true;
        }

        private void testTwoButton_Click(object sender, EventArgs e)
        {
            _presenter = new DataPresenter(this, false);
            _presenter.EmptySheetNames();
            _presenter.EmptyColumnNames();
            _presenter.DisplaySheetNames(fileName);
            sheetNameCombo.Enabled = true;
        }
    }
}
