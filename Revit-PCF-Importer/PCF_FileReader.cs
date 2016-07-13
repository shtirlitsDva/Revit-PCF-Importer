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
            readLines = System.IO.File.ReadAllLines(@"E:\CII\10\Ejby Vekslercentral_11-07-2016_12-18-30.pcf");
            return readLines;
        }
    }
}
