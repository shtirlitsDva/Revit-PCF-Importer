using System;
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


namespace Revit_PCF_Importer
{
    public static class PCF_Configuration
    {
        #region Export configuration
        public static IEnumerable<IGrouping<string, IGrouping<string, ElementSymbol>>>
            GroupSymbols(IList<ElementSymbol> symbolList)
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

        public static void ExportPipelinesElementsToExcel(IList<ElementSymbol> sourceList) 
            
        {
            //Grouped in pipelines
            var pipelineGroups = GroupByPipeline(sourceList);
            //Grouped in element types
            var typeGroups = GroupByElementType(sourceList);

            xel.Application excel = new xel.Application();
            if (null == excel)
            {
                Util.ErrorMsg("Failed to get or start Excel.");
            }
            excel.Visible = true;

            xel.Workbook workbook = excel.Workbooks.Add(Missing.Value);
            xel.Worksheet worksheet;
            worksheet = excel.ActiveSheet as xel.Worksheet; //New worksheet
            worksheet.Name = "Pipes and fittings"; //Name the created worksheet

            worksheet.Columns.ColumnWidth = 20;
            worksheet.Cells[1, 1] = "PIPELINE-REFERENCE"; //First column header
            //worksheet.Cells[1, 2] = "PCF Keyword ->"; //Second column header
            worksheet.Range["A1", Util.GetColumnName(typeGroups.Count()+1) + "1"].Font.Bold = true;

            //List of items that are handled by the PipeType and not instance placing
            //Remember to update this list!!!
            IList<string> typesHandledByPipeType = new List<string>();
            typesHandledByPipeType.Add("ELBOW");
            typesHandledByPipeType.Add("TEE");
            typesHandledByPipeType.Add("REDUCER-CONCENTRIC");
            typesHandledByPipeType.Add("REDUCER-ECCENTRIC");

            

            //Export the PCF keywords
            int row = 1, col = 1;

            //Iterate pipline references
            foreach (IGrouping<string, ElementSymbol> pipeLine in pipelineGroups)
            {
                //Write the pipeline values to EXCEL
                row++; //Increment row
                //worksheet.Cells[row, 1] = "PIPELINE-REFERENCE";
                worksheet.Cells[row, 1] = pipeLine.Key;
                worksheet.Range["A" + row, "A" + row].Font.Bold = true; //Make pipelines bold
            }
            //Write the top level keywords
            foreach (IGrouping<string, ElementSymbol> type in typeGroups)
            {
                col++; //Increment col
                worksheet.Cells[1, col] = type.Key;
                if (MyExtensions.ContainsString(typesHandledByPipeType, type.Key))
                {
                    for (row = 2; row <= pipelineGroups.Count()+1; row++ )
                    {
                        worksheet.Cells[row, col] = "Handled by PipeType";
                    }
                }
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
                    Regex rgx = new Regex("[^a-zA-Z0-9 -]");
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

            String[] excelSheetNames = new String[dt.Rows.Count];
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

                //Pipes and Fittings
                var pipesAndFittingsConf =
                    (from DataTable dataTable in dataTables
                        where string.Equals(dataTable.TableName, "Pipes and fittings")
                        select dataTable).FirstOrDefault();

                //query is using the variables in the loop to query the dataset
                EnumerableRowCollection<string> query = from value in pipesAndFittingsConf.AsEnumerable()
                    where value.Field<string>(0) == es.PipelineReference
                    select value.Field<string>(es.ElementType);

                string familyAndType = query.FirstOrDefault().ToString();

                FilteredElementCollector collector = new FilteredElementCollector(PCF_Importer_form._doc);

                ElementParameterFilter filter = Filter.ParameterValueFilter(familyAndType, BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM);

                LogicalOrFilter classFilter = Filter.FamSymbolsAndPipeTypes();

                Element familySymbol = collector.WherePasses(classFilter).WherePasses(filter).FirstOrDefault();

                if (es.ElementType == "PIPE") es.PipeType = (PipeType) familySymbol;
                else es.FamilySymbol = (FamilySymbol) familySymbol;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

        }

        #endregion

    }
}
