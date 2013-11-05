using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TwinArch.SMRT_MVPLibrary.Interfaces
{
    public interface ISMRTModel : ISMRTBaseInterface
    {
        List<string> GetSheetNames(string fileName);
        List<string> GetColumnNames(string fileName, string sheetName);
    }
}
