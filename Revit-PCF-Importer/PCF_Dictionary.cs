using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;

namespace Revit_PCF_Importer
{
    public class PCF_Dictionary : IParseKeywords
    {
        private readonly IKeywordProcessor _keywordProcessor;
        private Dictionary<string, Func<ElementSymbol, Result>> _dictionary;

        public PCF_Dictionary(IKeywordProcessor keywordProcessor)
        {
            _keywordProcessor = keywordProcessor;
            _dictionary = CreateDictionary();
        }

        public Dictionary<string, Func<ElementSymbol, Result>> CreateDictionary()
        {
            var dictionary = new Dictionary<string, Func<ElementSymbol, Result>>
            {
                {"ISOGEN-FILES", _keywordProcessor.ISOGEN_FILES},
                {"UNITS-BORE", _keywordProcessor.UNITS_BORE},
                {"UNITS-CO-ORDS", _keywordProcessor.UNITS_CO_ORDS},
                {"UNITS-WEIGHT", _keywordProcessor.UNITS_WEIGHT},
                {"UNITS-BOLT-DIA", _keywordProcessor.ELEMENT_TYPE_NOT_IMPLEMENTED},
                {"UNITS-BOLT-LENGTH", _keywordProcessor.ELEMENT_TYPE_NOT_IMPLEMENTED},
                {"UNITS-WEIGHT-LENGTH", _keywordProcessor.ELEMENT_TYPE_NOT_IMPLEMENTED},
                {"PIPELINE-REFERENCE", _keywordProcessor.PIPELINE_REFERENCE},
            };
            return dictionary;
        }

        public Result ProcessLevel1Keywords(ElementSymbol elementSymbol)
        {
            if (_dictionary.ContainsKey(elementSymbol.ElementType))
            {
                Result result = _dictionary[elementSymbol.ElementType].Invoke(elementSymbol);
                return result;
            }
            else
            {
                _keywordProcessor.ELEMENT_TYPE_NOT_IMPLEMENTED(elementSymbol);
                return Result.Failed;
            }
        }
    }
}
