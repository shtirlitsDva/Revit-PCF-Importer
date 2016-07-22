using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using PCF_Functions;

namespace Revit_PCF_Importer
{
    public class PCF_Dictionary : IParseKeywords
    {
        private readonly IKeywordProcessor _keywordProcessor;
        private Dictionary<string, Func<ElementSymbol, Result>> _topLevelDictionary;
        private Dictionary<string, Func<ElementSymbol, string, Result>> _elementLevelDictionary;

        public PCF_Dictionary(IKeywordProcessor keywordProcessor)
        {
            _keywordProcessor = keywordProcessor;
            _topLevelDictionary = CreateTopLevelDictionary();
            _elementLevelDictionary = CreateElementLevelDictionary();
        }

        public Dictionary<string, Func<ElementSymbol, Result>> CreateTopLevelDictionary()
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
                {"PIPE", _keywordProcessor.PIPE}
            };
            return dictionary;
        }

        public Dictionary<string, Func<ElementSymbol, string, Result>> CreateElementLevelDictionary()
        {
            var dictionary = new Dictionary<string, Func<ElementSymbol, string, Result>>
            {
                {"END-POINT", _keywordProcessor.END_POINT}
            };
            return dictionary;
        }

        public Result ProcessTopLevelKeywords(ElementSymbol elementSymbol)
        {
            if (_topLevelDictionary.ContainsKey(elementSymbol.ElementType))
            {
                Result result = _topLevelDictionary[elementSymbol.ElementType].Invoke(elementSymbol);
                return result;
            }
            else
            {
                _keywordProcessor.ELEMENT_TYPE_NOT_IMPLEMENTED(elementSymbol);
                return Result.Failed;
            }
        }

        public Result ProcessElementLevelKeywords(ElementSymbol elementSymbol, string line)
        {
            if (_elementLevelDictionary.ContainsKey(Parser.GetElementKeyword(line)))
            {
                Result result = _elementLevelDictionary[Parser.GetElementKeyword(line)].Invoke(elementSymbol, line);
                return result;
            }
            else
            {
                _keywordProcessor.ELEMENT_ATTRIBUTE_NOT_IMPLEMENTED(elementSymbol, line);
                return Result.Failed;
            }
        }
    }
}
