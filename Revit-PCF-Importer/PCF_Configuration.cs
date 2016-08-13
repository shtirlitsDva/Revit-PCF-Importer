using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using Autodesk.Revit.UI;
using BuildingCoder;
using xel = Microsoft.Office.Interop.Excel;

namespace Revit_PCF_Importer
{
    class PCF_Configuration
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

        public void ExportPipelinesElementsToExcel(
            IEnumerable<IGrouping<string, IGrouping<string, ElementSymbol>>> source)
        {
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
            worksheet.Cells[1, 1] = "PCF Keyword"; //First column header
            worksheet.Cells[1, 2] = "Family: Type"; //Second column header
            worksheet.Range["A1", "B2"].Font.Bold = true; //Make headers bold

            //Export the PCF keywords
            int row = 1;//, col = 2;

            var grouped = source;

            //Iterate pipline references
            foreach (IGrouping<string, IGrouping<string, ElementSymbol>> pipelineGroup in grouped)
            {
                //Ignore PRE-PIPELINE and MATERIALS
                if (string.Equals(pipelineGroup.Key, "PRE-PIPELINE") || string.Equals(pipelineGroup.Key, "MATERIALS")) continue;
                //Write the pipeline values to EXCEL
                row++; //Increment row
                worksheet.Cells[row, 1] = "PIPELINE-REFERENCE";
                worksheet.Cells[row, 2] = pipelineGroup.Key;
                //Write the top level keywords
                foreach (IGrouping<string, ElementSymbol> type in pipelineGroup)
                {
                    row++; //Increment row
                    worksheet.Cells[row, 1] = type.Key;
                }
            }
        }
    }
}
