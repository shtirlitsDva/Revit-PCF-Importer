using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using Autodesk.Revit.UI;
using BuildingCoder;
using PCF_Functions;
using xel = Microsoft.Office.Interop.Excel;


namespace Revit_PCF_Importer
{
    public static class PCF_Configuration
    {
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
            worksheet.Name = "PCF Configuration"; //Name the created worksheet

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
    }
}
