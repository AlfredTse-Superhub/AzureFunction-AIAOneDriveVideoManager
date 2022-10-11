using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OneDriveVideoManager.Models
{
    public class Checker
    {
        public string Email { get; set; }

        public List<Recording> Videos { get; set; }

        public string ListName { get; set; }
    }
}
