#region Namespaces
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
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
        //Declare the element collector
        public ElementSymbol[] ElementCollector;
        public ElementSymbol CurElementSymbol;
        public StringCollection[] SourceData;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            //Read the input file
            FileReader fileReader = new FileReader();
            string[] readFile = fileReader.ReadFile();

            StringCollection resultList;

            //Iteration counter
            int iterationCounter = -1;
            //Element index
            int curElementIndex = 0;
            
            foreach (string line in readFile)
            {
                //Count iterations
                iterationCounter++;

                //Logic test for Type or Property
                if (!line.StartsWith("    "))
                {
                    curElementIndex = iterationCounter;
                    CurElementSymbol = new ElementSymbol();
                }

                Idea: Parse the whole file and mark the position of pipeline references.

                ////Execute keyword handling
                ////Declare a StringCollection to hold the matches
                //StringCollection resultList = new StringCollection();

                ////Define a Regex to parse the input
                //Regex parseWords = new Regex(@"(\S+)");

                ////Define a Match to handle the results from Regex
                //Match match = parseWords.Match(line);

                ////Add every match from Regex to the StringCollection
                //while (match.Success)
                //{
                //    //Only add the result if it is not a white space or null
                //    if (!string.IsNullOrEmpty(match.Value)) resultList.Add(match.Value);
                //    match = match.NextMatch();
                //}
                ////Separate the keyword and the rest of words from the results
                //string keyword = resultList[0];
                ////Remove the keyword from the results
                //resultList.RemoveAt(0);
                //////Parse keywords
                ////var returns = dictionary.ParseKeywords(keyword, resultList);
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
