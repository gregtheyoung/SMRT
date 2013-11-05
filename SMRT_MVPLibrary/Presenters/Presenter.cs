using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TwinArch.SMRT_MVPLibrary.Interfaces;
using TwinArch.SMRT_MVPLibrary.Models;

namespace TwinArch.SMRT_MVPLibrary.Presenters
{
    public class Presenter<T> where T : ISMRTBaseInterface
    {
        protected ISMRTModel Model { get; private set; }
        protected T View { get; private set; }

        static Presenter()
        {
        }

        public Presenter(T view, bool useAutomation)
        {
            if (useAutomation)
                Model = new ExcelAutomationModel();
            else
                Model = new OLEDBModel();
            View = view;
        }
    }
}
