using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Revit_PCF_Importer
{
    public class ElementSymbol
    {
        public string ElementType { get; set; } //The element type.
        public StringCollection SourceData { get; set; } //This contains the raw data read from file
        public string PipelineReference { get; set; } //This contains the pipeline reference that was read
    }
}
