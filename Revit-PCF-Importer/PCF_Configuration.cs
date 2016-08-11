using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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


    }
}
