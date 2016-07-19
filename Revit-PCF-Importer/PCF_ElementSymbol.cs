using System;
using System.Collections;
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
        public StringCollection SourceData { get; set; } = new StringCollection();//This contains the raw data read from file
        public string PipelineReference { get; set; } = "PRE-PIPELINE"; //This contains the pipeline reference that was read
        public int Position { get; set; } //Contains the position of the element in file
        public int DefinitionLengthInLines { get; set; } //Contains the number of lines that hold the element data in PCF file
    }

    public class ElementCollection
    {
        public IList<ElementSymbol> Elements { get; set; } = new List<ElementSymbol>(); //The array to hold all the elements
        public IList<int> Position { get; set; } = new List<int>();//Holds the keywords line number in the file
    }
}
