using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace LinkChecker
{
    class ParserFacade
    {
        ParserFactory Factory = null;
        BaseParser parser = null;
        public ParserFacade()
        {
            Factory = new ParserFactory();
        }

        public void ParseFile(string fileUrl)
        {
            ParserType type = GetExtention(fileUrl);
            parser = Factory.GetParser(type.ToString());
            parser.FileUrl = fileUrl;
            parser.Parse();
        }

        public List<FileLink> GetParseResult()
        {
            List<FileLink> results = new List<FileLink>();
            Dictionary<string, BaseParser> parsers = Factory.GetAllParsers();
            foreach (BaseParser parser in parsers.Values)
            {
                results.AddRange(parser.Results);
            }
            return results;
        }

        private ParserType GetExtention(string fileUrl)
        {
            string strFileType = Path.GetExtension(fileUrl).ToLower();
            switch (strFileType)
            {
                case ".docx":
                    return ParserType.WORD;
                case ".pptx":
                    return ParserType.PPT;
                case ".mht":
                case ".aspx":
                    return ParserType.HTML;
                default:
                    return ParserType.INVALID;
            }
        }
    }
}
