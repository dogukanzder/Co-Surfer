using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoSurfer
{
    internal class Site
    {
        public string Title { get; set; }
        public string Url { get; set; }
        public string AutoWriteLocation { get; set; }
        public string SaveLocation { get; set; }
        public bool AutoStart { get; set; }
        public List<PathItem> Paths { get; set; }
    }
}
