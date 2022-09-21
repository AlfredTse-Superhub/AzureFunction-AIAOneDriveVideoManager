using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OneDriveVideoManager2.Models
{
    internal class AADGroup
    {
        public string GroupId { get; set; }

        public string GroupName { get; set; }

        public IList<DirectoryObject> MemberList { get; set; }
    }
}
