using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntegrityImporter
{
    public class Folder
    {
        public string Name { get; set; }
        public List<Folder> SubFolders { set; get; }

        public List<Requirement> Requirements { get; set; } 
    }
}
