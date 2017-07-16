using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Globalization;
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
using pif = Revit_PCF_Importer.PCF_Importer_form;

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
        Result GENERAL(ElementSymbol elementSymbol);
        Result FLANGE(ElementSymbol elementSymbol);

        //Element level keywords
        Result ELEMENT_ATTRIBUTE_NOT_IMPLEMENTED(ElementSymbol elementSymbol, string line);
        Result END_POINT(ElementSymbol elementSymbol, string line);
        Result CENTRE_POINT(ElementSymbol elementSymbol, string line);
        Result BRANCH1_POINT(ElementSymbol elementSymbol, string line);
        Result ANGLE(ElementSymbol elementSymbol, string line);
        Result MATERIAL_IDENTIFIER(ElementSymbol elementSymbol, string line);
        Result DESCRIPTION(ElementSymbol elementSymbol, string line);
        Result UCI(ElementSymbol elementSymbol, string line);
        Result SKEY(ElementSymbol elementSymbol, string line);
    }

    public class KeywordProcessor : IKeywordProcessor
    {
        //private Document doc = PCFImport.doc; //This code to expose doc to class, because I don't want to pass it to each method in the chain
        //                                      //See http://forums.autodesk.com/t5/revit-api/accessing-the-document-from-c-form-externalcommanddata-issue/td-p/4773407

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

        /// <summary>
        /// General handles all usual piping elements with no special requirements.
        /// </summary>
        /// <param name="elementSymbol"></param>
        /// <returns></returns>
        public Result GENERAL(ElementSymbol elementSymbol)
        {
            foreach (string line in elementSymbol.SourceData)
            {
                Result result = pif.PcfDict.ProcessElementLevelKeywords(elementSymbol, line);
                if (Result.Succeeded == result) continue;
                if (Result.Failed == result) return result;
            }
            return Result.Succeeded;
        }

        public Result FLANGE(ElementSymbol elementSymbol)
        {
            foreach (string line in elementSymbol.SourceData)
            {
                Result result = pif.PcfDict.ProcessElementLevelKeywords(elementSymbol, line);
                if (Result.Succeeded == result) continue;
                if (Result.Failed == result) return result;
            }

            elementSymbol.CentrePoint.Xyz = Util.Midpoint(elementSymbol.EndPoint1.Xyz, elementSymbol.EndPoint2.Xyz);
            elementSymbol.CentrePoint.Initialized = true;

            return Result.Succeeded;
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

            if (!elementSymbol.EndPoint1.Initialized)
            {
                elementSymbol.EndPoint1.Xyz = Parser.ParseXyz(endPointLine);
                elementSymbol.EndPoint1.Diameter = Parser.ParseDiameter(endPointLine);
                elementSymbol.EndPoint1.Initialized = true;
                return Result.Succeeded;
            }
            if (!elementSymbol.EndPoint2.Initialized)
            {
                elementSymbol.EndPoint2.Xyz = Parser.ParseXyz(endPointLine);
                elementSymbol.EndPoint2.Diameter = Parser.ParseDiameter(endPointLine);
                elementSymbol.EndPoint2.Initialized = true;
                return Result.Succeeded;
            }
            //The rest of line is ignored for now
            Util.ErrorMsg("Element at line number " + elementSymbol.Position + " has more than two END-POINTS, which is not allowed!");
            return Result.Failed;
        }

        public Result CENTRE_POINT(ElementSymbol elementSymbol, string line)
        {
            StringCollection endPointLine = Parser.GetRestOfTheLineInStringCollection(line);

            elementSymbol.CentrePoint.Xyz = Parser.ParseXyz(endPointLine);

            return Result.Succeeded;
        }

        public Result BRANCH1_POINT(ElementSymbol elementSymbol, string line)
        {
            StringCollection endPointLine = Parser.GetRestOfTheLineInStringCollection(line);

            elementSymbol.Branch1Point.Xyz = Parser.ParseXyz(endPointLine);
            elementSymbol.Branch1Point.Diameter = Parser.ParseDiameter(endPointLine);
            return Result.Succeeded;
        }

        public Result ANGLE(ElementSymbol elementSymbol, string line)
        {
            double angle = double.Parse(Parser.GetRestOfTheLine(line), CultureInfo.InvariantCulture);

            elementSymbol.Angle = angle/100;

            return Result.Succeeded;
        }

        public Result MATERIAL_IDENTIFIER(ElementSymbol elementSymbol, string line)
        {
            int matId = int.Parse(Parser.GetRestOfTheLine(line));
            elementSymbol.MaterialIdentifier = matId;
            return Result.Succeeded;
        }

        public Result DESCRIPTION(ElementSymbol elementSymbol, string line)
        {
            string description = Parser.GetRestOfTheLine(line);
            elementSymbol.MaterialDescription = description;
            return Result.Succeeded;
        }

        public Result UCI(ElementSymbol elementSymbol, string line)
        {
            string uci = Parser.GetRestOfTheLine(line);
            elementSymbol.UCI = uci;
            //Guid guid = new Guid(uci);
            //elementSymbol.guid = guid;
            return Result.Succeeded;
        }

        public Result SKEY(ElementSymbol elementSymbol, string line)
        {
            string skey = Parser.GetRestOfTheLine(line);
            elementSymbol.Skey = skey;
            return Result.Succeeded;
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
