using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TwinArch.SMRT_MVPLibrary.Interfaces
{
    public interface ISMRTMainView : ISMRTBase
    {
        List<string> SheetNames { set; }
        Dictionary<string, string> ColumnNames { set; }
        String AlertMessage { set; }
    }
}
