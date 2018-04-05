using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPOExporter
{
    //The config file data model that retrieves URL and stored/saved location
    public class SPOSite
    {
        public string URL { get; set; }

        public string savedDir { get; set; } 
    }
}
