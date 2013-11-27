using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TwinArch.SMRT_MVPLibrary
{
    public class TwitterUserInfo
    {
        public string Name { get; set; }
        public string Location { get; set; }
        public int NumberOfFollowers { get; set; }
        public int NumberFollowing { get; set; }
        public string Description { get; set; }
    }
}
