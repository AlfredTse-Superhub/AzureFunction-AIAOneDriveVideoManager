using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OneDriveVideoManager.Models
{
    public class FunctionRunLog
    {
        public string Id { get; set; }

        public string FunctionName { get; set; }

        public string Details { get; set; }

        public int TotalRecords { get; set; } = 0;

        public int UpdatedRecords { get; set; } = 0;

        public string Status { get; set; }

        public string LastStep { get; set; }
    }
}
