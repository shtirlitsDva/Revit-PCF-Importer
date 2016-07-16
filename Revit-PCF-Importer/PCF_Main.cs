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
                //Execute keyword handling
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
                //Separate the keyword and the rest of words from the results
                string keyword = resultList[0];
                //Remove the keyword from the results
                resultList.RemoveAt(0);
                //Instantiate the PCF_Dictionary mega complex refactoring switch solution which doesn't really work
                PCF_Dictionary dictionary = new PCF_Dictionary(new KeywordProcessor());
                dictionary.ParseKeywords(keyword, resultList);
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
