using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Revit_PCF_Importer
{
    public class PCF_Dictionary : IParseKeywords
    {
        private readonly IKeywordProcessor _keywordProcessor;
        private Dictionary<string, Func<string, StringCollection>> _dictionary;

        public PCF_Dictionary(IKeywordProcessor keywordProcessor)
        {
            _keywordProcessor = keywordProcessor;
            _dictionary = CreateDictionary();
        }

        public Dictionary<string, Func<string, StringCollection>> CreateDictionary()
        {
            var dictionary = new Dictionary<string, Func<string, StringCollection>>
            {
                {"ISOGEN-FILES", _keywordProcessor.ISOGEN_FILES(string keyword, StringCollection)},
                {"UNITS-BORE", _keywordProcessor.UNITS_BORE()},
                {"UNITS-CO-ORDS", _keywordProcessor.UNITS_CO_ORDS()}
            };
            return dictionary;
        }

        public string ParseKeywords(string keyword, StringCollection results)
        {
            if (_dictionary.ContainsKey(keyword))
            {
                return _dictionary[keyword].Invoke(_keywordProcessor.ISOGEN_FILES(results));
            }
            throw new Exception("Keyword not implemented!");
        }
    }
}
