using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LinkChecker
{
    public enum ParserType
    {
        WORD,
        PPT,
        HTML,
        INVALID
    }

    class ParserFactory
    {
        Dictionary<string, BaseParser> parsers = new Dictionary<string, BaseParser>();

        public ParserFactory()
        {
            BaseParser wordParser = new WordParser();
            parsers.Add(ParserType.WORD.ToString(), wordParser);
            BaseParser pptParser = new PowerPointParser();
            parsers.Add(ParserType.PPT.ToString(), pptParser);
            BaseParser htmlParser = new HtmlParser();
            parsers.Add(ParserType.HTML.ToString(), htmlParser);
        }

        public BaseParser GetParser(string fileType)
        {
            return parsers[fileType];
        }

        public Dictionary<string, BaseParser> GetAllParsers()
        {
            return parsers;
        }
    }
}
