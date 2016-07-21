using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using BuildingCoder;
using PCF_Functions;
using iv = PCF_Functions.InputVars;

namespace Revit_PCF_Importer
{
    public interface IParseKeywords
    {
        Result ProcessLevel1Keywords(ElementSymbol elementSymbol);
    }

    public interface IKeywordProcessor
    {
        Result ISOGEN_FILES(ElementSymbol elementSymbol);
        Result UNITS_BORE(ElementSymbol elementSymbol);
        Result UNITS_CO_ORDS(ElementSymbol elementSymbol);
        Result UNITS_WEIGHT(ElementSymbol elementSymbol);

    }

    public class KeywordProcessor : IKeywordProcessor
    {
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
            }
            if (string.Equals(value, "INCH"))
            {
                iv.UNITS_BORE_MM = false;
                iv.UNITS_BORE_INCH = true;
            }
            if (string.Equals(value, "INCH-SIXTEENTH"))
            {
                Util.ErrorMsg("Value INCH-SIXTEENTH for UNITS-BORE not implemented!\n" +
                              "Please specify either MM or INCH.");
                throw new Exception("Value not implemented!");
            }
            return Result.Succeeded;
        }

        public Result UNITS_CO_ORDS(ElementSymbol elementSymbol)
        {
            string value = Parser.GetRestOfTheLine(elementSymbol.SourceData[0]);
            iv.UNITS_CO_ORDS = value;
            if (string.Equals(value, "MM"))
            {
                iv.UNITS_CO_ORDS_MM = true;
                iv.UNITS_CO_ORDS_INCH = false;
            }
            if (string.Equals(value, "INCH"))
            {
                iv.UNITS_CO_ORDS_MM = false;
                iv.UNITS_CO_ORDS_INCH = true;
            }
            return Result.Succeeded;
        }

        public Result UNITS_WEIGHT(ElementSymbol elementSymbol)
        {
            string value = Parser.GetRestOfTheLine(elementSymbol.SourceData[0]);
            iv.UNITS_WEIGHT = value;
            if (string.Equals(value, "KGS"))
            {
                iv.UNITS_WEIGHT_KGS = true;
                iv.UNITS_WEIGHT_LBS = false;
            }
            if (string.Equals(value, "LBS"))
            {
                iv.UNITS_WEIGHT_KGS = false;
                iv.UNITS_WEIGHT_LBS = true;
            }
            return Result.Succeeded;
        }
    }
}
