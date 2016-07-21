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
        Result ProcessLevel1Keywords(ElementSymbol elementSymbol);
    }

    public interface IKeywordProcessor
    {
        Result ELEMENT_TYPE_NOT_IMPLEMENTED(ElementSymbol elementSymbol);
        Result ISOGEN_FILES(ElementSymbol elementSymbol);
        Result UNITS_BORE(ElementSymbol elementSymbol);
        Result UNITS_CO_ORDS(ElementSymbol elementSymbol);
        Result UNITS_WEIGHT(ElementSymbol elementSymbol);
        Result PIPELINE_REFERENCE(ElementSymbol elementSymbol);
        Result PIPE(ElementSymbol elementSymbol);
    }

    public class KeywordProcessor : IKeywordProcessor
    {
        private Document doc = PCFImport.doc; //This code to expose doc to class, because I don't want to pass it to each method in the chain
        //See http://forums.autodesk.com/t5/revit-api/accessing-the-document-from-c-form-externalcommanddata-issue/td-p/4773407

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
                //Instantiate collector
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                //Get the elements
                collector.OfClass(typeof(PipingSystemType));
                //Select correct systemType
                PipingSystemType sQuery = (from PipingSystemType st in collector
                                           where string.Equals(st.Abbreviation, elementSymbol.PipelineReference)
                                           select st).FirstOrDefault();
                elementSymbol.PipingSystemType = sQuery;
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

            IEnumerable<StringCollection> query = from string l in source
                                                  where string.Equals(Parser.GetElementKeyword(l), "END-POINT")
                                                  select Parser.GetRestOfTheLineInStringCollection(l);

            double X = double.Parse(query.FirstOrDefault()[0]);
            double Y = double.Parse(query.FirstOrDefault()[1]);
            double Z = double.Parse(query.FirstOrDefault()[2]);

            elementSymbol.EndPoint1.Xyz = new XYZ(X,Y,Z);

            return Result.Failed;
        }
    }

    public class HelperMethods
    {
        public static void ValueNotImplemented(ElementSymbol elementSymbol, string value)
        {
            throw new Exception("Value " + value + " for " + elementSymbol.ElementType + "  not implemented!\nSee program documentation for supported values.");
        }
    }
}
