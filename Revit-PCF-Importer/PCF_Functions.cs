using System;
using System.Data;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Linq;
using MoreLinq;
using System.Text;
using System.Globalization;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using BuildingCoder;
using Revit_PCF_Importer;
using iv = PCF_Functions.InputVars;
using pdef = PCF_Functions.ParameterDefinition;
using plst = PCF_Functions.ParameterList;
using pif = Revit_PCF_Importer.PCF_Importer_form;

namespace PCF_Functions
{
    public class InputVars
    {
        #region Execution

        //Used for "global variables".
        //File I/O
        public static string OutputDirectoryFilePath;
        public static string ExcelSheet = "COMP";

        //Execution control
        public static bool ConfigureAll = true;
        public static double DiameterLimit = 0;

        //PCF File Header (preamble) control
        public static string UNITS_BORE = "MM";
        public static bool UNITS_BORE_MM = true;
        public static bool UNITS_BORE_INCH = false;

        public static string UNITS_CO_ORDS = "MM";
        public static bool UNITS_CO_ORDS_MM = true;
        public static bool UNITS_CO_ORDS_INCH = false;

        public static string UNITS_WEIGHT = "KGS";
        public static bool UNITS_WEIGHT_KGS = true;
        public static bool UNITS_WEIGHT_LBS = false;

        public static string UNITS_WEIGHT_LENGTH = "METER";
        public static bool UNITS_WEIGHT_LENGTH_METER = true;
        //public static bool UNITS_WEIGHT_LENGTH_INCH = false; OBSOLETE
        public static bool UNITS_WEIGHT_LENGTH_FEET = false;

        #endregion Execution

        #region Filters

        //Filters
        public static string SysAbbr = "FVF";
        public static BuiltInParameter SysAbbrParam = BuiltInParameter.RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM;
        public static string PipelineGroupParameterName = "System Abbreviation";

        #endregion Filters
    }

    public class Composer
    {
        #region Preamble

        //PCF Preamble composition
        readonly StringBuilder sbPreamble = new StringBuilder();

        public StringBuilder PreambleComposer()
        {
            sbPreamble.Append("ISOGEN-FILES ISOGEN.FLS");
            sbPreamble.AppendLine();
            sbPreamble.Append("UNITS-BORE " + InputVars.UNITS_BORE);
            sbPreamble.AppendLine();
            sbPreamble.Append("UNITS-CO-ORDS " + InputVars.UNITS_CO_ORDS);
            sbPreamble.AppendLine();
            sbPreamble.Append("UNITS-WEIGHT " + InputVars.UNITS_WEIGHT);
            sbPreamble.AppendLine();
            sbPreamble.Append("UNITS-BOLT-DIA MM");
            sbPreamble.AppendLine();
            sbPreamble.Append("UNITS-BOLT-LENGTH MM");
            sbPreamble.AppendLine();
            sbPreamble.Append("UNITS-WEIGHT-LENGTH " + InputVars.UNITS_WEIGHT_LENGTH);
            sbPreamble.AppendLine();
            return sbPreamble;
        }

        #endregion

        #region Materials section

        StringBuilder sbMaterials = new StringBuilder();
        IEnumerable<IGrouping<string, Element>> materialGroups = null;
        int groupNumber;

        public StringBuilder MaterialsSection(IEnumerable<IGrouping<string, Element>> elementGroups)
        {
            materialGroups = elementGroups;
            sbMaterials.Append("MATERIALS");
            foreach (IGrouping<string, Element> group in materialGroups)
            {
                groupNumber++;
                sbMaterials.AppendLine();
                sbMaterials.Append("MATERIAL-IDENTIFIER " + groupNumber);
                sbMaterials.AppendLine();
                sbMaterials.Append("    DESCRIPTION " + group.Key);
            }
            return sbMaterials;
        }

        #endregion

        #region CII export writer

        StringBuilder sbCII;
        private Document doc;
        private string key;

        public StringBuilder CIIWriter(Document document, string systemAbbreviation)
        {
            doc = document;
            key = systemAbbreviation;
            sbCII = new StringBuilder();
            //Handle CII export parameters
            //Instantiate collector
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            //Get the elements
            collector.OfClass(typeof (PipingSystemType));
            //Select correct systemType
            PipingSystemType sQuery = (from PipingSystemType st in collector
                where string.Equals(st.Abbreviation, key)
                select st).FirstOrDefault();

            var query = from p in new plst().ListParametersAll
                where string.Equals(p.Domain, "PIPL") && string.Equals(p.ExportingTo, "CII")
                select p;

            foreach (pdef p in query.ToList())
            {
                if (string.IsNullOrEmpty(sQuery.get_Parameter(p.Guid).AsString())) continue;
                sbCII.Append("    ");
                sbCII.Append(p.Keyword);
                sbCII.Append(" ");
                sbCII.Append(sQuery.get_Parameter(p.Guid).AsString());
                sbCII.AppendLine();
            }

            return sbCII;
        }

        #endregion

        #region ELEM parameter writer

        private StringBuilder sbElemParameters;
        private Element element;

        public StringBuilder ElemParameterWriter(Element passedElement)
        {
            sbElemParameters = new StringBuilder();
            element = passedElement;
            var pQuery = from p in new plst().ListParametersAll
                where !string.IsNullOrEmpty(p.Keyword) && string.Equals(p.Domain, "ELEM")
                select p;

            foreach (pdef p in pQuery)
            {
                //Check for parameter's storage type (can be Int for select few parameters)
                int sT = (int) element.get_Parameter(p.Guid).StorageType;

                if (sT == 1)
                {
                    //Check if the parameter contains anything
                    if (string.IsNullOrEmpty(element.get_Parameter(p.Guid).AsInteger().ToString())) continue;
                    sbElemParameters.Append("    " + p.Keyword + " ");
                    sbElemParameters.Append(element.get_Parameter(p.Guid).AsInteger());
                }
                else if (sT == 3)
                {
                    //Check if the parameter contains anything
                    if (string.IsNullOrEmpty(element.get_Parameter(p.Guid).AsString())) continue;
                    sbElemParameters.Append("    " + p.Keyword + " ");
                    sbElemParameters.Append(element.get_Parameter(p.Guid).AsString());
                }
                sbElemParameters.AppendLine();
            }
            return sbElemParameters;
        }

        #endregion
    }

    public class Filter
    {
        public static ElementParameterFilter ParameterValueFilter(string valueQualifier, BuiltInParameter parameterName)
        {
            BuiltInParameter testParam = parameterName;
            ParameterValueProvider pvp = new ParameterValueProvider(new ElementId((int) testParam));
            FilterStringRuleEvaluator str = new FilterStringContains();
            FilterStringRule paramFr = new FilterStringRule(pvp, str, valueQualifier, false);
            ElementParameterFilter epf = new ElementParameterFilter(paramFr);
            return epf;
        }

        public static LogicalOrFilter FamSymbolsAndPipeTypes()
        {
            BuiltInCategory[] bics = new BuiltInCategory[]
            {
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_PipeFitting,
            };

            IList<ElementFilter> a = new List<ElementFilter>(bics.Count());

            foreach (BuiltInCategory bic in bics) a.Add(new ElementCategoryFilter(bic));

            LogicalOrFilter categoryFilter = new LogicalOrFilter(a);

            LogicalAndFilter familySymbolFilter = new LogicalAndFilter(categoryFilter,
                new ElementClassFilter(typeof(FamilySymbol)));

            IList<ElementFilter> b = new List<ElementFilter>();

            b.Add(new ElementClassFilter(typeof(PipeType)));

            b.Add(familySymbolFilter);

            LogicalOrFilter classFilter = new LogicalOrFilter(b);

            return classFilter;
        }
    }

    public class FilterDiameterLimit
    {
        private Element element;
        private bool diameterLimitBool;
        private double diameterLimit;

        /// <summary>
        /// Tests the diameter of the pipe or primary connector of element against the diameter limit set in the interface.
        /// </summary>
        /// <param name="passedElement"></param>
        /// <returns>True if diameter is larger than limit and false if smaller.</returns>
        public bool FilterDL(Element passedElement)
        {
            element = passedElement;
            diameterLimit = iv.DiameterLimit;
            diameterLimitBool = true;
            double testedDiameter = 0;
            switch (element.Category.Id.IntegerValue)
            {
                case (int) BuiltInCategory.OST_PipeCurves:
                    if (iv.UNITS_BORE_MM)
                        testedDiameter = double.Parse(Conversion.PipeSizeToMm(((MEPCurve) element).Diameter/2));
                    else if (iv.UNITS_BORE_INCH)
                        testedDiameter = double.Parse(Conversion.PipeSizeToInch(((MEPCurve) element).Diameter/2));

                    if (testedDiameter <= diameterLimit) diameterLimitBool = false;

                    break;

                case (int) BuiltInCategory.OST_PipeFitting:
                case (int) BuiltInCategory.OST_PipeAccessory:
                    //Cast the element passed to method to FamilyInstance
                    FamilyInstance familyInstance = (FamilyInstance) element;
                    //MEPModel of the elements is accessed
                    MEPModel mepmodel = familyInstance.MEPModel;
                    //Get connector set for the element
                    ConnectorSet connectorSet = mepmodel.ConnectorManager.Connectors;
                    //Declare a variable for 
                    Connector testedConnector = null;

                    if (connectorSet.IsEmpty) break;
                    if (connectorSet.Size == 1)
                        foreach (Connector connector in connectorSet) testedConnector = connector;
                    else
                        testedConnector = (from Connector connector in connectorSet
                            where connector.GetMEPConnectorInfo().IsPrimary
                            select connector).FirstOrDefault();

                    if (iv.UNITS_BORE_MM)
                        testedDiameter = double.Parse(Conversion.PipeSizeToMm(testedConnector.Radius));
                    else if (iv.UNITS_BORE_INCH)
                        testedDiameter = double.Parse(Conversion.PipeSizeToInch(testedConnector.Radius));

                    if (testedDiameter <= diameterLimit) diameterLimitBool = false;

                    break;
            }
            return diameterLimitBool;
        }
    }

    public class Conversion
    {
        const double _inch_to_mm = 25.4;
        const double _foot_to_mm = 12*_inch_to_mm;
        const double _foot_to_inch = 12;

        /// <summary>
        /// Return a string for a real number formatted to two decimal places.
        /// </summary>
        public static string RealString(double a)
        {
            //return a.ToString("0.##");
            return (Math.Truncate(a*100)/100).ToString("0.00", CultureInfo.GetCultureInfo("en-GB"));
        }

        /// <summary>
        /// Return a string for an XYZ point or vector with its coordinates converted from feet to millimetres and formatted to two decimal places.
        /// </summary>
        public static string PointStringMm(XYZ p)
        {
            return string.Format("{0:0.00} {1:0.00} {2:0.00}",
                RealString(p.X*_foot_to_mm),
                RealString(p.Y*_foot_to_mm),
                RealString(p.Z*_foot_to_mm));
        }

        public static string PointStringInch(XYZ p)
        {
            return string.Format("{0:0.00} {1:0.00} {2:0.00}",
                RealString(p.X*_foot_to_inch),
                RealString(p.Y*_foot_to_inch),
                RealString(p.Z*_foot_to_inch));
        }

        public static string PipeSizeToMm(double l)
        {
            return string.Format("{0}", Math.Round(l*2*_foot_to_mm));
        }

        public static string PipeSizeToInch(double l)
        {
            return string.Format("{0}", RealString(l*2*_foot_to_inch));
        }

        public static string AngleToPCF(double l)
        {
            return string.Format("{0}", l);
        }

        public static double RadianToDegree(double angle)
        {
            return angle*(180.0/Math.PI);
        }

        public static double DegreeToRadian(double angle)
        {
            return Math.PI * angle / 180.0;
        }
    }

    public class ScheduleCreator
    {
        private UIDocument _uiDoc;

        public Result CreateAllItemsSchedule(UIDocument uiDoc)
        {
            try
            {
                _uiDoc = uiDoc;
                Document doc = _uiDoc.Document;
                FilteredElementCollector sharedParameters = new FilteredElementCollector(doc);
                sharedParameters.OfClass(typeof (SharedParameterElement));

                #region Debug

                ////Debug
                //StringBuilder sbDev = new StringBuilder();
                //var list = new ParameterDefinition().ElementParametersAll;
                //int i = 0;

                //foreach (SharedParameterElement sp in sharedParameters)
                //{
                //    sbDev.Append(sp.GuidValue + "\n");
                //    sbDev.Append(list[i].Guid.ToString() + "\n");
                //    i++;
                //    if (i == list.Count) break;
                //}
                ////sbDev.Append( + "\n");
                //// Clear the output file
                //File.WriteAllBytes(InputVars.OutputDirectoryFilePath + "\\Dev.pcf", new byte[0]);

                //// Write to output file
                //using (StreamWriter w = File.AppendText(InputVars.OutputDirectoryFilePath + "\\Dev.pcf"))
                //{
                //    w.Write(sbDev);
                //    w.Close();
                //}

                #endregion

                Transaction t = new Transaction(doc, "Create items schedules");
                t.Start();

                #region Schedule ALL elements

                ViewSchedule schedAll = ViewSchedule.CreateSchedule(doc, ElementId.InvalidElementId,
                    ElementId.InvalidElementId);
                schedAll.Name = "PCF - ALL Elements";
                schedAll.Definition.IsItemized = false;

                IList<SchedulableField> schFields = schedAll.Definition.GetSchedulableFields();

                foreach (SchedulableField schField in schFields)
                {
                    if (schField.GetName(doc) != "Family and Type") continue;
                    ScheduleField field = schedAll.Definition.AddField(schField);
                    ScheduleSortGroupField sortGroupField = new ScheduleSortGroupField(field.FieldId);
                    schedAll.Definition.AddSortGroupField(sortGroupField);
                }

                string curUsage = "U";
                string curDomain = "ELEM";
                var query = from p in new plst().ListParametersAll
                    where p.Usage == curUsage && p.Domain == curDomain
                    select p;

                foreach (pdef pDef in query.ToList())
                {
                    SharedParameterElement parameter = (from SharedParameterElement param in sharedParameters
                        where param.GuidValue.CompareTo(pDef.Guid) == 0
                        select param).First();
                    SchedulableField queryField =
                        (from fld in schFields
                            where fld.ParameterId.IntegerValue == parameter.Id.IntegerValue
                            select fld).First();

                    ScheduleField field = schedAll.Definition.AddField(queryField);
                    if (pDef.Name != "PCF_ELEM_TYPE") continue;
                    ScheduleFilter filter = new ScheduleFilter(field.FieldId, ScheduleFilterType.HasParameter);
                    schedAll.Definition.AddFilter(filter);
                }

                #endregion

                #region Schedule FILTERED elements

                ViewSchedule schedFilter = ViewSchedule.CreateSchedule(doc, ElementId.InvalidElementId,
                    ElementId.InvalidElementId);
                schedFilter.Name = "PCF - Filtered Elements";
                schedFilter.Definition.IsItemized = false;

                schFields = schedFilter.Definition.GetSchedulableFields();

                foreach (SchedulableField schField in schFields)
                {
                    if (schField.GetName(doc) != "Family and Type") continue;
                    ScheduleField field = schedFilter.Definition.AddField(schField);
                    ScheduleSortGroupField sortGroupField = new ScheduleSortGroupField(field.FieldId);
                    schedFilter.Definition.AddSortGroupField(sortGroupField);
                }

                foreach (pdef pDef in query.ToList())
                {
                    SharedParameterElement parameter = (from SharedParameterElement param in sharedParameters
                        where param.GuidValue.CompareTo(pDef.Guid) == 0
                        select param).First();
                    SchedulableField queryField =
                        (from fld in schFields
                            where fld.ParameterId.IntegerValue == parameter.Id.IntegerValue
                            select fld).First();

                    ScheduleField field = schedFilter.Definition.AddField(queryField);
                    if (pDef.Name != "PCF_ELEM_TYPE") continue;
                    ScheduleFilter filter = new ScheduleFilter(field.FieldId, ScheduleFilterType.HasParameter);
                    schedFilter.Definition.AddFilter(filter);
                    filter = new ScheduleFilter(field.FieldId, ScheduleFilterType.NotEqual, "");
                    schedFilter.Definition.AddFilter(filter);
                }

                #endregion

                #region Schedule Pipelines

                ViewSchedule schedPipeline = ViewSchedule.CreateSchedule(doc,
                    new ElementId(BuiltInCategory.OST_PipingSystem), ElementId.InvalidElementId);
                schedPipeline.Name = "PCF - Pipelines";
                schedPipeline.Definition.IsItemized = false;

                schFields = schedPipeline.Definition.GetSchedulableFields();

                foreach (SchedulableField schField in schFields)
                {
                    if (schField.GetName(doc) != "Family and Type") continue;
                    ScheduleField field = schedPipeline.Definition.AddField(schField);
                    ScheduleSortGroupField sortGroupField = new ScheduleSortGroupField(field.FieldId);
                    schedPipeline.Definition.AddSortGroupField(sortGroupField);
                }

                curDomain = "PIPL";
                foreach (pdef pDef in query.ToList())
                {
                    SharedParameterElement parameter = (from SharedParameterElement param in sharedParameters
                        where param.GuidValue.CompareTo(pDef.Guid) == 0
                        select param).First();
                    SchedulableField queryField =
                        (from fld in schFields
                            where fld.ParameterId.IntegerValue == parameter.Id.IntegerValue
                            select fld).First();
                    schedPipeline.Definition.AddField(queryField);
                }

                #endregion

                t.Commit();

                sharedParameters.Dispose();

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                Util.InfoMsg(e.Message);
                return Result.Failed;
            }
        }
    }

    public class Parser
    {
        public static void CreateInitialElementList(ElementCollection collection, string[] source)
        {
            int iterationCounter = -1;

            //Holds current pipeline reference
            string curPipelineReference = "PRE-PIPELINE";

            foreach (string line in source)
            {
                //Count iterations
                iterationCounter++;

                //Logic test for Type or Property
                if (!line.StartsWith("    "))
                {
                    //Make a new Element
                    ElementSymbol CurElementSymbol = new ElementSymbol();
                    //Get the keyword from the parsed line
                    CurElementSymbol.ElementType = GetElementKeyword(line);
                    //Get the element position in the file
                    CurElementSymbol.Position = iterationCounter;
                    //Set the correct pipeline reference
                    //Add here if other top levet types needed that is on the same level as PIPELINE-REFERENCE
                    switch (CurElementSymbol.ElementType)
                    {
                        case "PIPELINE-REFERENCE":
                            curPipelineReference = GetRestOfTheLine(line);
                            break;
                        case "MATERIALS":
                            curPipelineReference = "MATERIALS";
                            break;
                    }
                    CurElementSymbol.PipelineReference = curPipelineReference;

                    //Get the PipingSystemType based on the reference
                    //Instantiate collector
                    FilteredElementCollector collector = new FilteredElementCollector(pif._doc);

                    ElementParameterFilter filter = Filter.ParameterValueFilter(curPipelineReference, BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM);
                    //Get the elements
                    PipingSystemType sQuery = collector.OfClass(typeof (PipingSystemType)).WherePasses(filter).Cast<PipingSystemType>().FirstOrDefault();

                    if (sQuery != null) CurElementSymbol.PipingSystemType = sQuery;

                    //Add the extracted element to the collection
                    collection.Elements.Add(CurElementSymbol);
                    collection.Position.Add(iterationCounter);
                }
            }
        }

        public static string GetElementKeyword(string line)
        {
            //Execute keyword handling
            //Define a Regex to parse the input
            Regex parseWords = new Regex(@"(\S+)");

            //Define a Match to handle the results from Regex
            Match match = parseWords.Match(line);

            //Separate the keyword and the rest of words from the results
            string keyword = match.Value;

            return keyword;
        }

        public static string GetRestOfTheLine(string line)
        {
            //Execute keyword handling
            //Declare a StringCollection to hold the matches
            StringCollection resultList = new StringCollection();

            //Define a Regex to parse the input
            Regex parseWords = new Regex(@"(\S+)");

            //Define a Match to handle the results from Regex
            Match match = parseWords.Match(line);

            //Add every match from Regex to the StringCollection
            while (match.Success)
            {
                //Only add the result if it is not a white space or null
                if (!string.IsNullOrEmpty(match.Value)) resultList.Add(match.Value);
                match = match.NextMatch();
            }
            //Remove the keyword from the results
            resultList.RemoveAt(0);

            string restOfTheLine = string.Empty;

            //Concat the StringCollection to a string
            if (resultList.Count == 1)
            {
                restOfTheLine = resultList[0];
            }

            if (resultList.Count > 1)
            {
                string[] strArray = new string[resultList.Count];
                resultList.CopyTo(strArray, 0);
                restOfTheLine = string.Join(" ", strArray);
            }

            return restOfTheLine;
        }

        public static StringCollection GetRestOfTheLineInStringCollection(string line)
        {
            //Execute keyword handling
            //Declare a StringCollection to hold the matches
            StringCollection resultList = new StringCollection();

            //Define a Regex to parse the input
            Regex parseWords = new Regex(@"(\S+)");

            //Define a Match to handle the results from Regex
            Match match = parseWords.Match(line);

            //Add every match from Regex to the StringCollection
            while (match.Success)
            {
                //Only add the result if it is not a white space or null
                if (!string.IsNullOrEmpty(match.Value)) resultList.Add(match.Value);
                match = match.NextMatch();
            }
            //Remove the keyword from the results
            resultList.RemoveAt(0);

            return resultList;
        }

        /// <summary>
        /// Compares the position value of elements in list and calculates number of lines in the element definition.
        /// </summary>
        /// <param name="collection"></param>
        /// <param name="source"></param>
        public static void IndexElementDefinitions(ElementCollection collection, string[] source)
        {
            for (int idx = 0; idx < collection.Elements.Count; idx++)
            {
                //Handle last element
                if (collection.Elements.Count == idx + 1)
                {
                    int lastIndex = source.Length - 1;
                    collection.Elements[idx].DefinitionLengthInLines = lastIndex - collection.Elements[idx].Position;
                    continue;
                }
                //Handle all other elements
                int differenceInPosition = collection.Elements[idx + 1].Position - collection.Elements[idx].Position - 1;
                collection.Elements[idx].DefinitionLengthInLines = differenceInPosition;
            }
        }

        public static void ExtractElementDefinition(ElementCollection collection, string[] source)
        {
            foreach (ElementSymbol e in collection.Elements)
            {
                //Handle all elements
                int curPosition = e.Position;
                int defLength = e.DefinitionLengthInLines;
                StringCollection collectedLines = new StringCollection();

                //Iterate over the lines, very important to make sure that the defLength is defined correctly
                for (int i = 0; i <= defLength; i++)
                {
                    collectedLines.Add(source[curPosition]);
                    curPosition++;
                }
                //Push the extracted lines to the element
                e.SourceData = collectedLines;
            }
        }

        public static XYZ ParseXyz(StringCollection endPointLine)
        {
            double X = double.Parse(endPointLine[0], CultureInfo.InvariantCulture);
            double Y = double.Parse(endPointLine[1], CultureInfo.InvariantCulture);
            double Z = double.Parse(endPointLine[2], CultureInfo.InvariantCulture);

            if (iv.UNITS_CO_ORDS_MM)
            {
                X = Util.MmToFoot(X);
                Y = Util.MmToFoot(Y);
                Z = Util.MmToFoot(Z);
            }

            if (iv.UNITS_CO_ORDS_INCH)
            {
                X = Util.InchToFoot(X);
                Y = Util.InchToFoot(Y);
                Z = Util.InchToFoot(Z);
            }

            XYZ xyz = new XYZ(X, Y, Z);

            return xyz;
        }

        public static double ParseDiameter(StringCollection endPointLine)
        {
            double diameter = double.Parse(endPointLine[3], CultureInfo.InvariantCulture);

            if (iv.UNITS_BORE_MM) diameter = Util.MmToFoot(diameter);

            if (iv.UNITS_BORE_INCH) diameter = Util.InchToFoot(diameter);

            return diameter;
        }
    }

    public class CreatorHelper
    {
        public static FilteredElementCollector GetElementsWithConnectors(Document doc)
        {
            // what categories of family instances
            // are we interested in?
            // From here: http://thebuildingcoder.typepad.com/blog/2010/06/retrieve-mep-elements-and-connectors.html

            BuiltInCategory[] bics = new BuiltInCategory[]
            {
                //BuiltInCategory.OST_CableTray,
                //BuiltInCategory.OST_CableTrayFitting,
                //BuiltInCategory.OST_Conduit,
                //BuiltInCategory.OST_ConduitFitting,
                //BuiltInCategory.OST_DuctCurves,
                //BuiltInCategory.OST_DuctFitting,
                //BuiltInCategory.OST_DuctTerminal,
                //BuiltInCategory.OST_ElectricalEquipment,
                //BuiltInCategory.OST_ElectricalFixtures,
                //BuiltInCategory.OST_LightingDevices,
                //BuiltInCategory.OST_LightingFixtures,
                //BuiltInCategory.OST_MechanicalEquipment,
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_PipeFitting,
                //BuiltInCategory.OST_PlumbingFixtures,
                //BuiltInCategory.OST_SpecialityEquipment,
                //BuiltInCategory.OST_Sprinklers,
                //BuiltInCategory.OST_Wire
            };

            IList<ElementFilter> a = new List<ElementFilter>(bics.Count());

            foreach (BuiltInCategory bic in bics) a.Add(new ElementCategoryFilter(bic));

            LogicalOrFilter categoryFilter = new LogicalOrFilter(a);

            LogicalAndFilter familyInstanceFilter = new LogicalAndFilter(categoryFilter, new ElementClassFilter(typeof(FamilyInstance)));

            //IList<ElementFilter> b = new List<ElementFilter>(6);
            IList<ElementFilter> b = new List<ElementFilter>();

            //b.Add(new ElementClassFilter(typeof(CableTray)));
            //b.Add(new ElementClassFilter(typeof(Conduit)));
            //b.Add(new ElementClassFilter(typeof(Duct)));
            b.Add(new ElementClassFilter(typeof(Pipe)));

            b.Add(familyInstanceFilter);

            LogicalOrFilter classFilter = new LogicalOrFilter(b);

            FilteredElementCollector collector = new FilteredElementCollector(doc);

            collector.WherePasses(classFilter);

            return collector;
        }

        public static ConnectorSet GetConnectorSet(Element e)
        {
            ConnectorSet connectors = null;

            if (e is FamilyInstance)
            {
                MEPModel m = ((FamilyInstance)e).MEPModel;
                if (null != m && null != m.ConnectorManager) connectors = m.ConnectorManager.Connectors;
            }

            else if (e is Wire) connectors = ((Wire)e).ConnectorManager.Connectors;

            else
            {
                Debug.Assert(e.GetType().IsSubclassOf(typeof(MEPCurve)), 
                  "expected all candidate connector provider "
                  + "elements to be either family instances or "
                  + "derived from MEPCurve");

                if (e is MEPCurve) connectors = ((MEPCurve)e).ConnectorManager.Connectors;
            }
            return connectors;
        }

        public static HashSet<Connector> GetAllPipeConnectors()
        {
            //Get a list of all pipes created in project
            //Consider adding a filter to only work on pipes in the same pipeline
            HashSet<MEPCurve> query = (from ElementSymbol es in PCFImport.ExtractedElementCollection.Elements
                                    where string.Equals("PIPE", es.ElementType) && es.CreatedElement != null //Make sure no null elements are passed around, sigh....
                                    select (MEPCurve)es.CreatedElement).ToHashSet();

            //Collect all pipe connectors present im project while filtering for end connector types
            HashSet<Connector> allPipeConnectors = query
                .Select(mepCurve => mepCurve.ConnectorManager.Connectors)
                .SelectMany(conSet => (from Connector c in conSet
                                       where (int)c.ConnectorType == 1
                                       select c)).ToHashSet();
            return allPipeConnectors;
        }

        public static IList<Connector> GetALLConnectors(Document doc)
        {
            return (from e in GetElementsWithConnectors(doc) from Connector c in GetConnectorSet(e) select c).ToList();
        }

        public static Pipe CreateDummyPipe(XYZ pointToConnect, XYZ directionPoint, PointInSpace endInstance, ElementSymbol elementSymbol)
        {
            Pipe pipe = null;

            ElementId pipeTypeId = elementSymbol.PipeType.Id;

            //Collect levels and select one level
            FilteredElementCollector levelCollector = new FilteredElementCollector(PCFImport.doc);
            ElementClassFilter levelFilter = new ElementClassFilter(typeof(Level));
            ElementId levelId = levelCollector.WherePasses(levelFilter).FirstElementId();
            
            //Create vector to define pipe
            XYZ pipeDir = pointToConnect - directionPoint;
            XYZ helperPoint1 = pointToConnect.Add(pipeDir.Multiply(2));
            //Create pipe
            pipe = Pipe.Create(PCFImport.doc, elementSymbol.PipingSystemType.Id, pipeTypeId, levelId,
                pointToConnect, helperPoint1);
            //Set pipe diameter
            Parameter parameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
            parameter.Set(endInstance.Diameter);

            return pipe;
        }

        public static Connector MatchConnector(XYZ pointToMatch, Pipe pipe)
        {
            Connector connector = null;
            connector = (from Connector c in pipe.ConnectorManager.Connectors
                  where Util.IsEqual(pointToMatch, c.Origin)
                  select c).FirstOrDefault();

            return connector;
        }

        public static Connector MatchConnector(XYZ pointToMatch, ConnectorSet conSet)
        {
            Connector connector = null;
            connector = (from Connector c in conSet
                         where Util.IsEqual(pointToMatch, c.Origin)
                         select c).FirstOrDefault();

            return connector;
        }

        public static Connector GetSecondaryConnector(ConnectorSet conSet)
        {
            Connector connector = null;
            connector = (from Connector c in conSet
                         where c.GetMEPConnectorInfo().IsSecondary.Equals(true)
                         select c).FirstOrDefault();

            return connector;
        }

        /// <summary>
        /// Return a 3D view from the given document.
        /// </summary>
        public static View3D Get3DView(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);

            collector.OfClass(typeof(View3D));

            foreach (View3D v in collector)
            {
                // skip view templates here because they
                // are invisible in project browsers:

                if (v != null && !v.IsTemplate && v.Name == "{3D}")
                {
                    return v;
                }
            }
            return null;
        }
    }

    public class Helper
    {
        /// <summary>
        /// This method is used to place an adaptive family which helps in debugging
        /// </summary>
        /// <param name="typeName"></param>
        /// <param name="p1"></param>
        /// <param name="p2"></param>
        /// <returns></returns>
        public static FamilyInstance PlaceAdaptiveMarkerLine(string typeName, XYZ p1, XYZ p2)
        {
            //Get the symbol
            ElementParameterFilter filter = Filter.ParameterValueFilter("Marker Line: " + typeName,
                BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM); //Hardcoded until implements
            FamilySymbol markerSymbol =
                new FilteredElementCollector(PCFImport.doc).WherePasses(filter)
                    .Cast<FamilySymbol>()
                    .FirstOrDefault();
            // Create a new instance of an adaptive component family
            FamilyInstance instance = AdaptiveComponentInstanceUtils.CreateAdaptiveComponentInstance(PCFImport.doc,
                markerSymbol);
            // Get the placement points of this instance
            IList<ElementId> placePointIds = new List<ElementId>();
            placePointIds = AdaptiveComponentInstanceUtils.GetInstancePlacementPointElementRefIds(instance);
            // Set the position of each placement point
            ReferencePoint point1 = PCFImport.doc.GetElement(placePointIds[0]) as ReferencePoint;
            point1.Position = p1;
            ReferencePoint point2 = PCFImport.doc.GetElement(placePointIds[1]) as ReferencePoint;
            point2.Position = p2;

            return instance;
        }
    }

    public static class MyExtensions
    {
        /// <summary>
        /// Extension method which searches a list of strings to determine if it contains a specific string.
        /// http://www.dotnetperls.com/list-contains
        /// </summary>
        /// <param name="list"></param>
        /// <param name="item"></param>
        /// <returns></returns>
        public static bool ContainsString(this IList<string> list, string item)
        {
            for (int i = 0; i < list.Count; i++)
            {
                if (list[i] == item)
                {
                    return true;
                }
            }
            return false;
        }
    }
}