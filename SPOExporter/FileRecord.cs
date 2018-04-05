using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPOExporter
{
    //The file mapping data model in the file
    public class FileRecord
    {
        public string FilePath { get; set; }
        public string LastModifiedDate { get; set; }
    }
}
