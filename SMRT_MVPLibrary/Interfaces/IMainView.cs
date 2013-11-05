using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TwinArch.SMRT_MVPLibrary.Interfaces
{
    public interface ISMRTMainView : ISMRTBaseInterface
    {
        List<string> SheetNames { set; }
        List<string> ColumnNames { set; }
    }
}
