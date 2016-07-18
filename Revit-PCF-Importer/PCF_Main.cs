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
using PCF_Functions;

#endregion

namespace Revit_PCF_Importer
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class PCFImport : IExternalCommand
    {
        //Declare the element collector
        public ElementCollection ExtractedElementCollection;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            ExtractedElementCollection = new ElementCollection();

            //Read the input file
            FileReader fileReader = new FileReader();
            string[] readFile = fileReader.ReadFile();

            //Iteration counter
            int iterationCounter = -1;
            //Parser
            Parser parser = new Parser();
            //Holds current pipeline reference
            string curPipelineReference = "PRE-PIPELINE";
            
            //This loop collects all top-level element strings and creates ElementSymbols with data
            foreach (string line in readFile)
            {
                //Count iterations
                iterationCounter++;

                //Logic test for Type or Property
                if (!line.StartsWith("    "))
                {
                    //Make a new Element
                    ElementSymbol CurElementSymbol = new ElementSymbol();
                    //Get the keyword from the parsed line
                    CurElementSymbol.ElementType = parser.GetElementKeyword(line);
                    //Get the element position in the file
                    CurElementSymbol.Position = iterationCounter;
                    switch (CurElementSymbol.ElementType)
                    {
                        case "PIPELINE-REFERENCE":
                        case "MATERIALS":
                            curPipelineReference = parser.GetRestOfTheLine(line);
                            break;
                    }
                    CurElementSymbol.PipelineReference = curPipelineReference;

                    //Add the extracted element to the collection
                    ExtractedElementCollection.Elements.Add(CurElementSymbol);
                    ExtractedElementCollection.Position.Add(iterationCounter);
                }
            }

            //This loop compares all element symbols and gets the amount of line for their definition
            //Then it extracts the lines to a property --- NEEDS TO BE IMPLEMENTED!
            for (int idx = 0; idx < ExtractedElementCollection.Elements.Count; idx++)
            {
                //Handle last element
                if (ExtractedElementCollection.Elements.Count == idx + 1)
                {
                    int lastIndex = readFile.Length - 1;
                    ExtractedElementCollection.Elements[idx].DefinitionLengthInLines = ExtractedElementCollection.Elements[idx].Position - lastIndex;
                    continue;
                }
                
                int differenceInPosition = ExtractedElementCollection.Elements[idx + 1].Position - ExtractedElementCollection.Elements[idx].Position - 1;
                ExtractedElementCollection.Elements[idx].DefinitionLengthInLines = differenceInPosition;

                //Implement extraction of lines
            }

            //!!!Handle the last element in the listarray


            //using (Transaction tx = new Transaction(doc))
            //{
            //    tx.Start("Transaction Name");
            //    tx.Commit();
            //}

            return Result.Succeeded;
        }
    }
}
