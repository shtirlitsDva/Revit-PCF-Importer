using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using BuildingCoder;
using PCF_Functions;
using iv = PCF_Functions.InputVars;
using hm = Revit_PCF_Importer.HelperMethods;

namespace Revit_PCF_Importer
{
    public interface IParseKeywords
    {
        Result ProcessTopLevelKeywords(ElementSymbol elementSymbol);
        Result ProcessElementLevelKeywords(ElementSymbol elementSymbol, string line);
    }

    public interface IKeywordProcessor
    {
        //Top level keywords
        Result ELEMENT_TYPE_NOT_IMPLEMENTED(ElementSymbol elementSymbol);
        Result ISOGEN_FILES(ElementSymbol elementSymbol);
        Result UNITS_BORE(ElementSymbol elementSymbol);
        Result UNITS_CO_ORDS(ElementSymbol elementSymbol);
        Result UNITS_WEIGHT(ElementSymbol elementSymbol);
        Result PIPELINE_REFERENCE(ElementSymbol elementSymbol);
        Result PIPE(ElementSymbol elementSymbol);

        //Element level keywords
        Result ELEMENT_ATTRIBUTE_NOT_IMPLEMENTED(ElementSymbol elementSymbol, string line);
        Result END_POINT(ElementSymbol elementSymbol, string line);
    }

    public class KeywordProcessor : IKeywordProcessor
    {
        private Document doc = PCFImport.doc; //This code to expose doc to class, because I don't want to pass it to each method in the chain
                                              //See http://forums.autodesk.com/t5/revit-api/accessing-the-document-from-c-form-externalcommanddata-issue/td-p/4773407

        private PCF_Dictionary pcfDict = new PCF_Dictionary(new KeywordProcessor());

        #region Top level keywords
        public Result ELEMENT_TYPE_NOT_IMPLEMENTED(ElementSymbol elementSymbol)
        {
            return Result.Succeeded;
        }

        public Result ISOGEN_FILES(ElementSymbol elementSymbol)
        {
            return Result.Succeeded;
        }

        public Result UNITS_BORE(ElementSymbol elementSymbol)
        {
            string value = Parser.GetRestOfTheLine(elementSymbol.SourceData[0]);

            iv.UNITS_BORE = value;
            if (string.Equals(value, "MM"))
            {
                iv.UNITS_BORE_MM = true;
                iv.UNITS_BORE_INCH = false;
                return Result.Succeeded;
            }
            if (string.Equals(value, "INCH"))
            {
                iv.UNITS_BORE_MM = false;
                iv.UNITS_BORE_INCH = true;
                return Result.Succeeded;
            }
            hm.ValueNotImplemented(elementSymbol, value);
            return Result.Failed;
        }

        public Result UNITS_CO_ORDS(ElementSymbol elementSymbol)
        {
            string value = Parser.GetRestOfTheLine(elementSymbol.SourceData[0]);
            iv.UNITS_CO_ORDS = value;
            if (string.Equals(value, "MM"))
            {
                iv.UNITS_CO_ORDS_MM = true;
                iv.UNITS_CO_ORDS_INCH = false;
                return Result.Succeeded;
            }
            if (string.Equals(value, "INCH"))
            {
                iv.UNITS_CO_ORDS_MM = false;
                iv.UNITS_CO_ORDS_INCH = true;
                return Result.Succeeded;
            }
            hm.ValueNotImplemented(elementSymbol, value);
            return Result.Failed;
        }

        public Result UNITS_WEIGHT(ElementSymbol elementSymbol)
        {
            string value = Parser.GetRestOfTheLine(elementSymbol.SourceData[0]);
            iv.UNITS_WEIGHT = value;
            if (string.Equals(value, "KGS"))
            {
                iv.UNITS_WEIGHT_KGS = true;
                iv.UNITS_WEIGHT_LBS = false;
                return Result.Succeeded;
            }
            if (string.Equals(value, "LBS"))
            {
                iv.UNITS_WEIGHT_KGS = false;
                iv.UNITS_WEIGHT_LBS = true;
                return Result.Succeeded;
            }
            hm.ValueNotImplemented(elementSymbol, value);
            return Result.Failed;
        }

        public Result PIPELINE_REFERENCE(ElementSymbol elementSymbol)
        {
            try
            {
                ////Instantiate collector
                //FilteredElementCollector collector = new FilteredElementCollector(doc);
                ////Get the elements
                //collector.OfClass(typeof(PipingSystemType));
                ////Select correct systemType
                //PipingSystemType sQuery = (from PipingSystemType st in collector
                //                           where string.Equals(st.Abbreviation, elementSymbol.PipelineReference)
                //                           select st).FirstOrDefault();
                //elementSymbol.PipingSystemType = sQuery;
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                Util.ErrorMsg(e.Message);
                return Result.Failed;
            }
        }

        public Result PIPE(ElementSymbol elementSymbol)
        {
            StringCollection source = elementSymbol.SourceData;

            foreach (string s in source)
            {
                Result result = pcfDict.ProcessElementLevelKeywords(elementSymbol, s);
                return result;
            }
            return Result.Failed;
        }
        #endregion

        #region Element level keywords
        public Result ELEMENT_ATTRIBUTE_NOT_IMPLEMENTED(ElementSymbol elementSymbol, string line)
        {
            return Result.Succeeded;
        }

        public Result END_POINT(ElementSymbol elementSymbol, string line)
        {
            StringCollection endPointLine = Parser.GetRestOfTheLineInStringCollection(line);

            double X = double.Parse(endPointLine[0]);
            double Y = double.Parse(endPointLine[1]);
            double Z = double.Parse(endPointLine[2]);

            if (!elementSymbol.EndPoint1.Initialized)
            {
                elementSymbol.EndPoint1.Xyz = new XYZ(X, Y, Z);
                elementSymbol.EndPoint1.Diameter = double.Parse(endPointLine[3]);
                elementSymbol.EndPoint1.Initialized = true;
                return Result.Succeeded;
            }
            if (!elementSymbol.EndPoint2.Initialized)
            {
                elementSymbol.EndPoint2.Xyz = new XYZ(X, Y, Z);
                elementSymbol.EndPoint2.Diameter = double.Parse(endPointLine[3]);
                elementSymbol.EndPoint2.Initialized = true;
                return Result.Succeeded;
            }
            //The rest of line is ignored for now
            Util.ErrorMsg("Element at line number " + elementSymbol.Position + " has more than two END-POINTS, which is not allowed!");
            return Result.Failed;
        }
        #endregion
    }

    public class HelperMethods
    {
        public static void ValueNotImplemented(ElementSymbol elementSymbol, string value)
        {
            throw new Exception("Value " + value + " for " + elementSymbol.ElementType + "  not implemented!\nSee program documentation for supported values.");
        }
    }
}
