using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Revit_PCF_Importer
{
    class FileReader
    {
        string[] readLines;

        public string[] ReadFile()
        {
            readLines = System.IO.File.ReadAllLines(@"G:\CII\12\Ejby Vekslercentral_18-07-2016_22-59-32.pcf");
            return readLines;
        }
    }
}
