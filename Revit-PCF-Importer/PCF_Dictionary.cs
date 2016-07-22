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
                {"END-POINT", _keywordProcessor.END_POINT},
                {"CENTRE-POINT", _keywordProcessor.CENTRE_POINT },
                {"ANGLE", _keywordProcessor.ANGLE },
                {"MATERIAL-IDENTIFIER", _keywordProcessor.MATERIAL_IDENTIFIER },
                {"DESCRIPTION", _keywordProcessor.DESCRIPTION },
                {"UCI", _keywordProcessor.UCI },
                {"UNIQUE-COMPONENT-IDENTIFIER", _keywordProcessor.UCI }
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
                Result result = _keywordProcessor.ELEMENT_TYPE_NOT_IMPLEMENTED(elementSymbol);
                return result;
            }
        }

        public Result ProcessElementLevelKeywords(ElementSymbol elementSymbol, string line)
        {
            string keyword = Parser.GetElementKeyword(line);

            if (_elementLevelDictionary.ContainsKey(keyword))
            {
                Result result = _elementLevelDictionary[keyword].Invoke(elementSymbol, line);
                return result;
            }
            else
            {
                Result result = _keywordProcessor.ELEMENT_ATTRIBUTE_NOT_IMPLEMENTED(elementSymbol, line);
                return result;
            }
        }
    }

    public class PCF_Creator : ICreateElements
    {
        private readonly IProcessElements _processElements;
        private Dictionary<string, Func<ElementSymbol, Result>> _elementCreationDictionary;

        public PCF_Creator(IProcessElements processElements)
        {
            _processElements = processElements;
            _elementCreationDictionary = CreateElementCreationDictionary();
        }

        private Dictionary<string, Func<ElementSymbol, Result>> CreateElementCreationDictionary()
        {
            var dictionary = new Dictionary<string, Func<ElementSymbol, Result>>
            {
                {"PIPE", _processElements.PIPE }
            };
            return dictionary;
        }

        public Result SendElementsToCreation(ElementSymbol elementSymbol)
        {
            if (_elementCreationDictionary.ContainsKey(elementSymbol.ElementType))
            {
                Result result = _elementCreationDictionary[elementSymbol.ElementType].Invoke(elementSymbol);
                return result;
            }
            else
            {
                Result result = _processElements.ELEMENT_TYPE_NOT_IMPLEMENTED(elementSymbol);
                return result;
            }
        }
    }
}
