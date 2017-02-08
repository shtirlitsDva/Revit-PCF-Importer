using System;
using System.CodeDom;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using BuildingCoder;
using PCF_Functions;
using xel = Microsoft.Office.Interop.Excel;
using iv = PCF_Functions.InputVars;


namespace Revit_PCF_Importer
{
    public static class PCF_Configuration
    {
        #region Export configuration
        public static IEnumerable<IGrouping<string, IGrouping<string, ElementSymbol>>>
            GroupSymbolsByPipelineThenType(IList<ElementSymbol> symbolList)
        {
            //Nested groupings: https://msdn.microsoft.com/da-dk/library/bb545974.aspx
            //Group all elementSymbols by pipeline and element type
            var grouped = from ElementSymbol es in symbolList
                group es by es.PipelineReference
                into pipeLineGroup
                from elementTypeGroup in
                    (from ElementSymbol es in pipeLineGroup
                        group es by es.ElementType)
                group elementTypeGroup by pipeLineGroup.Key;

            return grouped;
        }

        public static IEnumerable<IGrouping<string, IGrouping<string, ElementSymbol>>> GroupSymbolsByTypeThenSkey(
            IList<ElementSymbol> symbolList)
        {
            var grouped = from ElementSymbol es in symbolList
                where !(
                    string.Equals(es.PipelineReference, "PRE-PIPELINE") ||
                    string.Equals(es.PipelineReference, "MATERIALS") ||
                    string.Equals(es.ElementType, "PIPELINE-REFERENCE")
                    )
                group es by es.ElementType
                into typeGroup
                from skeyGroup in
                    (from ElementSymbol es in typeGroup
                        group es by es.Skey)
                group skeyGroup by typeGroup.Key;

            return grouped;
        }

        public static IEnumerable<IGrouping<string, PointInSpace>> GroupEndPointsByDiameter(IList<ElementSymbol> symbolList)
        {
            var elementSymbols = from ElementSymbol es in symbolList
                where !(
                    string.Equals(es.PipelineReference, "PRE-PIPELINE") ||
                    string.Equals(es.PipelineReference, "MATERIALS") ||
                    string.Equals(es.ElementType, "PIPELINE-REFERENCE")
                    )
                select es;

            IList<PointInSpace> allPointsInSpace = new List<PointInSpace>();
            foreach (ElementSymbol es in elementSymbols)
            {
                if (es.EndPoint1.Diameter != 0) allPointsInSpace.Add(es.EndPoint1);
                if (es.EndPoint2.Diameter != 0) allPointsInSpace.Add(es.EndPoint2);
                if (es.CentrePoint.Diameter != 0) allPointsInSpace.Add(es.CentrePoint);
                if (es.CoOrds.Diameter != 0) allPointsInSpace.Add(es.CoOrds);
                if (es.Branch1Point.Diameter != 0) allPointsInSpace.Add(es.Branch1Point);
            }
            //Order list to make the diameters appear in order from small to large
            IList<PointInSpace> orderedList = allPointsInSpace.OrderBy(pis => pis.Diameter).ToList();

            IEnumerable<IGrouping<string, PointInSpace>> grouped;

            if (iv.UNITS_BORE_MM) grouped = from PointInSpace pis in orderedList group pis by Conversion.PipeSizeToMm(pis.Diameter/2);
            else grouped = from PointInSpace pis in orderedList group pis by Conversion.PipeSizeToInch(pis.Diameter/2);
            
            return grouped;
        }

        public static IEnumerable<IGrouping<string, ElementSymbol>> GroupByPipeline(IList<ElementSymbol> symbolList)
        {
            var grouped = from ElementSymbol es in symbolList
                where !(
                    string.Equals(es.PipelineReference, "PRE-PIPELINE") ||
                    string.Equals(es.PipelineReference, "MATERIALS")
                    )
                group es by es.PipelineReference;

            return grouped;
        }

        public static IEnumerable<IGrouping<string, ElementSymbol>> GroupByElementType(IList<ElementSymbol> symbolList)
        {
            var grouped = from ElementSymbol es in (from ElementSymbol es in symbolList
                where !(
                    string.Equals(es.PipelineReference, "PRE-PIPELINE") ||
                    string.Equals(es.PipelineReference, "MATERIALS") ||
                    string.Equals(es.ElementType, "PIPELINE-REFERENCE")
                    )
                select es)
                group es by es.ElementType;

            return grouped;
        }

        public static void ExportAllConfigurationToExcel(IList<ElementSymbol> sourceList)
        {
            //Grouped in pipelines
            var typeAndSkeyGroups = GroupSymbolsByTypeThenSkey(sourceList);
            //Grouped in element types
            var diameterGroups = GroupEndPointsByDiameter(sourceList);

            xel.Application excel = new xel.Application();
            if (null == excel)
            {
                Util.ErrorMsg("Failed to get or start Excel.");
            }
            excel.Visible = true;

            xel.Workbook workbook = excel.Workbooks.Add(Missing.Value);
            xel.Worksheet worksheet;
            worksheet = excel.ActiveSheet as xel.Worksheet; //New worksheet
            worksheet.Name = "All pipelines"; //Name the created worksheet

            worksheet.Columns.ColumnWidth = 20;
            worksheet.Cells[1, 1] = "Type and skey"; //First column header
            worksheet.Cells[1, 2] = "All sizes"; //Second column header
            worksheet.Range["A1", Util.GetColumnName(diameterGroups.Count()+2) + "1"].Font.Bold = true;

            //Export the PCF keywords
            int row = 1, col = 2;

            //Write diameters
            foreach (var diameter in diameterGroups)
            {
                //Write the pipeline values to EXCEL
                col++;
                worksheet.Cells[1, col] = diameter.Key;
            }
            //Write elements and skeys
            foreach (var type in typeAndSkeyGroups)
            {
                row++;
                worksheet.Cells[row, 1] = type.Key;
                worksheet.Range["A" + row, "A" + row].Font.Bold = true;
                foreach (var skey in type)
                {
                    if (!skey.Key.IsNullOrEmpty()) continue; //<--- Why does ! make it work???
                    row++;
                    worksheet.Cells[row, 1] = skey.Key;
                }
            }
        }

        public static void ExportByPipelineConfigurationToExcel(IList<ElementSymbol> sourceList)
        {
            try
            {
                //Group pipeline elements by pipeline ref, type and skey
                //http://stackoverflow.com/a/6133139/6073998 triple nested linq
                var grouped = from ElementSymbol es in sourceList
                    where !string.Equals(es.ElementType, "PIPELINE-REFERENCE")
                    group es by new {es.PipelineReference, es.ElementType, es.Skey}
                    into g2
                    group g2 by new {g2.Key.PipelineReference, g2.Key.ElementType}
                    into g1
                    group g1 by g1.Key.PipelineReference;

                //Init excel instance and create a workbook
                xel.Application excel = new xel.Application();
                if (null == excel)
                {
                    Util.ErrorMsg("Failed to get or start EXCEL.");
                    throw new Exception("Failed to get or start EXCEL");
                }
                excel.Visible = true;

                xel.Workbook workbook = excel.Workbooks.Add(Missing.Value);
                xel.Worksheet worksheet;

                //To create distinction between first iteration and others
                bool first = true;

                //Iterate groups and write out the data
                foreach (var pipeline in grouped)
                {
                    string pipelineRef = pipeline.Key;
                    if (first)
                    {
                        worksheet = excel.ActiveSheet as xel.Worksheet; //select default worksheet
                        worksheet.Name = pipelineRef; //Name the created worksheet

                        worksheet.Columns.ColumnWidth = 20;
                        worksheet.Cells[1, 1] = "Type and skey"; //First column header
                        worksheet.Cells[1, 2] = "All sizes"; //Second column header
                        first = false;
                    }
                    else
                    {
                        excel.Sheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        worksheet = excel.ActiveSheet as xel.Worksheet;
                        worksheet.Name = pipelineRef;

                        worksheet.Columns.ColumnWidth = 20;
                        worksheet.Cells[1, 1] = "Type and skey"; //First column header
                        worksheet.Cells[1, 2] = "All sizes"; //Second column header
                    }

                    //Manage the diameter grouping
                    var pipelineElements = (from ElementSymbol es in sourceList
                        where (
                            !string.Equals(es.ElementType, "PIPELINE-REFERENCE") &&
                            string.Equals(es.PipelineReference, pipelineRef)
                            )
                        select es).ToList();

                    var diameterGroups = GroupEndPointsByDiameter(pipelineElements);
                    //Make the top row bold now that we know number of sizes
                    worksheet.Range["A1", Util.GetColumnName(diameterGroups.Count() + 2) + "1"].Font.Bold = true;

                    //Export the PCF keywords
                    int row = 1, col = 2;

                    //Write diameters
                    foreach (var diameter in diameterGroups)
                    {
                        //Write the pipeline values to EXCEL
                        col++;
                        worksheet.Cells[1, col] = diameter.Key;
                    }
                    //Write elements and skeys
                    foreach (var type in pipeline)
                    {
                        row++;
                        worksheet.Cells[row, 1] = type.Key.ElementType;
                        worksheet.Range["A" + row, "A" + row].Font.Bold = true;
                        foreach (var skey in type)
                        {
                            if (!skey.Key.Skey.IsNullOrEmpty()) continue; //<--- Why does ! make it work???
                            row++;
                            worksheet.Cells[row, 1] = skey.Key.Skey;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw new Exception(e.Message);
            }
        }

        #endregion

        #region Import configuration

        //DataSet import is from here:
        //http://stackoverflow.com/a/18006593/6073998
        public static DataSet ImportExcelToDataSet(string fileName)
        {
            //On connection strings http://www.connectionstrings.com/excel/#p84
            string connectionString =
                string.Format(
                    "provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"",
                    fileName);

            DataSet data = new DataSet();

            foreach (string sheetName in GetExcelSheetNames(connectionString))
            {
                using (OleDbConnection con = new OleDbConnection(connectionString))
                {
                    var dataTable = new DataTable();
                    string query = string.Format("SELECT * FROM [{0}]", sheetName);
                    con.Open();
                    OleDbDataAdapter adapter = new OleDbDataAdapter(query, con);
                    adapter.Fill(dataTable);

                    //Remove ' and $ from sheetName
                    Regex rgx = new Regex("[^a-zA-Z0-9 _-]");
                    string tableName = rgx.Replace(sheetName, "");

                    dataTable.TableName = tableName;
                    data.Tables.Add(dataTable);
                }
            }

            if (data == null) Util.ErrorMsg("Data set is null");
            if (data.Tables.Count < 1) Util.ErrorMsg("Table count in DataSet is 0");
            
            return data;
        }

        static string[] GetExcelSheetNames(string connectionString)
        {
            OleDbConnection con = null;
            DataTable dt = null;
            con = new OleDbConnection(connectionString);
            con.Open();
            dt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            if (dt == null)
            {
                return null;
            }

            string[] excelSheetNames = new string[dt.Rows.Count];
            int i = 0;

            foreach (DataRow row in dt.Rows)
            {
                excelSheetNames[i] = row["TABLE_NAME"].ToString();
                i++;
            }

            return excelSheetNames;
        }

        public static void ExtractElementConfiguration(DataSet dataSet, ElementSymbol es)
        {
            try
            {
                DataTableCollection dataTables = dataSet.Tables;
                DataTable dataTable;

                //Handle all pipelines or separate configuration setting
                if (iv.ConfigureAll)
                {
                    dataTable = (from DataTable dt in dataTables
                                 where string.Equals(dt.TableName, "All pipelines")
                                 select dt).FirstOrDefault();
                }
                else
                {
                    dataTable = (from DataTable dt in dataTables
                                 where string.Equals(dt.TableName, es.PipelineReference)
                                 select dt).FirstOrDefault();
                }

                var who = es.ElementType;

                //query the element family and type is using the variables in the loop to query the dataset
                EnumerableRowCollection<string> query = from value in dataTable.AsEnumerable()
                    where value.Field<string>(0) == es.PipelineReference
                    select value.Field<string>(es.ElementType);
                string familyAndType = query.FirstOrDefault().ToString();
                FilteredElementCollector collector = new FilteredElementCollector(PCF_Importer_form._doc);
                ElementParameterFilter filter = Filter.ParameterValueFilter(familyAndType, BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM);
                LogicalOrFilter classFilter = Filter.FamSymbolsAndPipeTypes();
                Element familySymbol = collector.WherePasses(classFilter).WherePasses(filter).FirstOrDefault();

                if (es.ElementType == "PIPE") es.PipeType = (PipeType)familySymbol;
                else es.FamilySymbol = (FamilySymbol)familySymbol;

                //query the corresponding pipe family and type to add to the element symbol
                //This is because pipe type is needed to create certain fittings
                if (es.ElementType != "PIPE" || es.ElementType != "OLET") //Exclude olets -- they are handled next
                {
                    EnumerableRowCollection<string> queryPipeType = from value in dataTable.AsEnumerable()
                                                            where value.Field<string>(0) == es.PipelineReference
                                                            select value.Field<string>("PIPE");
                    string pipeTypeName = queryPipeType.FirstOrDefault();
                    FilteredElementCollector collectorPipeType = new FilteredElementCollector(PCF_Importer_form._doc);
                    ElementParameterFilter filterPipeTypeName = Filter.ParameterValueFilter(pipeTypeName, BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM);
                    Element pipeType = collectorPipeType.OfClass(typeof(PipeType)).WherePasses(filterPipeTypeName).FirstOrDefault();
                    es.PipeType = (PipeType)pipeType;
                }
                if (es.ElementType == "OLET") //Get the TAP pipetype for olets
                {
                    EnumerableRowCollection<string> queryPipeType = from value in dataTable.AsEnumerable()
                                                                    where value.Field<string>(0) == "Olet"
                                                                    select value.Field<string>("PIPE");
                    string pipeTypeName = queryPipeType.FirstOrDefault();
                    FilteredElementCollector collectorPipeType = new FilteredElementCollector(PCF_Importer_form._doc);
                    ElementParameterFilter filterPipeTypeName = Filter.ParameterValueFilter(pipeTypeName, BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM);
                    Element pipeType = collectorPipeType.OfClass(typeof(PipeType)).WherePasses(filterPipeTypeName).FirstOrDefault();
                    es.PipeType = (PipeType)pipeType;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }
        #endregion

    }
}
