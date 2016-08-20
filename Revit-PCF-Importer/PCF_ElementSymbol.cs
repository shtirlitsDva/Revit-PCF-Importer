using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;

namespace Revit_PCF_Importer
{
    public class ElementSymbol
    {
        public string ElementType { get; set; } //The element type.
        public StringCollection SourceData { get; set; } = new StringCollection();//This contains the raw data read from file
        public string PipelineReference { get; set; } = "PRE-PIPELINE"; //This contains the pipeline reference that was read
        public string Skey { get; set; } = string.Empty; //This contains the element SKEY from source
        public PipingSystemType PipingSystemType { get; set; } //Holds the PipeLineType for the element
        public FamilySymbol FamilySymbol { get; set; } //Holds the family symbol if appropriate
        public PipeType PipeType { get; set; } //Holds the pipe type if appropriate
        public int Position { get; set; } //Contains the position of the element in file
        public int DefinitionLengthInLines { get; set; } //Contains the number of lines that hold the element data in PCF file
        public PointInSpace EndPoint1 = new PointInSpace("END-POINT");
        public PointInSpace EndPoint2 = new PointInSpace("END-POINT");
        public PointInSpace CoOrds = new PointInSpace("CO-ORDS");
        public PointInSpace CentrePoint = new PointInSpace("CENTRE-POINT");
        public PointInSpace Branch1Point = new PointInSpace("BRANCH1-POINT");
        public int MaterialIdentifier { get; set; } = 0;
        public string MaterialDescription { get; set; } = string.Empty;
        public double Angle { get; set; } = -999;
        //public Guid guid { get; set; } = null;
        public string UCI { get; set; } = string.Empty;
        public Element CreatedElement { get; set; } = null; //Holds the element that was created based on the information stored in the object
        public Element DummyToDelete { get; set; } = null; //Holds any dummy element created which could not be deleted in the same transaction
        public bool IsReducing { get; set; } = false;
    }
    /// <summary>
    /// Holds the coordinate information read from file.
    /// </summary>
    public class PointInSpace
    {
        public string Keyword { get; set; } //Contains the keyword for the pointInSpace
        public XYZ Xyz { get; set; } = null;
        public double Diameter { get; set; } = new double();
        public string RestOfTheLine { get; set; } = string.Empty;
        public bool Initialized { get; set; } = false;

        public PointInSpace(string keyword)
        {
            Keyword = keyword;
        }

    }

    public class ElementCollection
    {
        public IList<ElementSymbol> Elements { get; set; } = new List<ElementSymbol>(); //The list to hold all the elements
        public IList<int> Position { get; set; } = new List<int>();//Holds the line number in the file
    }
}
