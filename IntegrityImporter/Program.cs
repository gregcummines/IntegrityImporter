using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntegrityImporter
{
    class Program
    {
        static void Main(string[] args)
        {
            var folder = args[0];
            var files = Directory.GetFiles(folder);
            foreach (var file in files)
            {
                var requirements = IntegrityImporter.ParseDocumentAsHtml(file);
            }
        }
    }
}
