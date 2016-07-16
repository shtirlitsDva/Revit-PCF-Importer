using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BuildingCoder;
using iv = PCF_Functions.InputVars;

namespace Revit_PCF_Importer
{
    public interface IParseKeywords
    {
        string ParseKeywords(string keyword, StringCollection results);
    }

    public interface IKeywordProcessor
    {
        void ISOGEN_FILES(StringCollection results);
        void UNITS_BORE(StringCollection results);
        void UNITS_CO_ORDS(StringCollection results);
    }

    public class KeywordProcessor : IKeywordProcessor
    {
        public void ISOGEN_FILES(StringCollection results)
        {
            //Do nothing!
        }

        public void UNITS_BORE(StringCollection results)
        {
            string value = results[1];
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
        }

        public void UNITS_CO_ORDS(StringCollection results)
        {
            string value = results[1];
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
        }
    }
}
