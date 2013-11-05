using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TwinArch.SMRT_MVPLibrary.Interfaces;

namespace TwinArch.SMRT_MVPLibrary.Presenters
{
    public class DataPresenter : Presenter<ISMRTMainView>
    {
        public DataPresenter(ISMRTMainView view, bool useAutomation)
            : base(view, useAutomation)
        {
        }

        public void DisplaySheetNames(string filename)
        {
            if (View != null) View.SheetNames = Model.GetSheetNames(filename);
        }

        public void DisplayColumnNames(string filename, string sheetname)
        {
            if (View != null) View.ColumnNames = Model.GetColumnNames(filename, sheetname);
        }

        public void EmptySheetNames()
        {
            if (View != null) View.SheetNames = null;
        }

        public void EmptyColumnNames()
        {
            if (View != null) View.ColumnNames = null;
        }


    }
}
