#region Namespaces
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

using iv = PCF_Functions.InputVars;
using BuildingCoder;
#endregion

namespace Revit_PCF_Importer
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class PCFImport : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            //Read the input file
            FileReader fileReader = new FileReader();
            string[] readFile = fileReader.ReadFile();

            foreach (string line in readFile)
            {
                if (!line.StartsWith("    "))
                {
                    //Execute type handling
                    //Declare a StringCollection to hold the matches
                    StringCollection resultList = new StringCollection();
                    //Define a Regex to parse the input
                    Regex parseWords = new Regex(@"(\w+|[-]|[.])*");
                    //Define a Match to handle the results from Regex
                    Match match = parseWords.Match(line);
                    //Add every match from Regex to the StringCollection
                    while (match.Success)
                    {
                        //Only add the result if it is not a white space or null
                        if (!string.IsNullOrEmpty(match.Value)) resultList.Add(match.Value);
                        match = match.NextMatch();
                    }

                    //Declare helper variables
                    string value;

                    switch (resultList[0])
                    {
                        case "ISOGEN-FILES":
                            break;

                        case "UNITS-BORE":
                            #region UNITS-BORE
                            value = resultList[1];
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
                            if (string.Equals(value,"INCH-SIXTEENTH"))
                            {
                                Util.ErrorMsg("Value INCH-SIXTEENTH for UNITS-BORE not implemented!\n" +
                                              "Please specify either MM or INCH.");
                            }
                            #endregion
                            break;
                        case "UNITS-CO-ORDS":
                            #region UNITS-CO-ORDS
                            value = resultList[1];
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
                            #endregion
                            break;
                    }

                }
                if (line.StartsWith("    "))
                {
                    //Execute attribute handling
                }
            }

            //using (Transaction tx = new Transaction(doc))
            //{
            //    tx.Start("Transaction Name");
            //    tx.Commit();
            //}

            return Result.Succeeded;
        }
    }
}
