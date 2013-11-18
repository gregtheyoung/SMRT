using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TwinArch.SMRT_MVPLibrary.Interfaces
{
    public interface ISMRTDoman : ISMRTBase, IDisposable
    {
        List<string> GetSheetNames(string fileName);
        Dictionary<string, string> GetColumnNames(string fileName, string sheetName);
        bool SplitURLs(string fileName, string sheetName, string urlColumnName);
    }
}
