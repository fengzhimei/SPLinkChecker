using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace LinkChecker
{
    abstract class BaseParser
    {
        public string FileUrl { get; set; }
        public List<FileLink> Results { get; set; }

        public BaseParser()
        {
            Results = new List<FileLink>();
        }

        public BaseParser(string FileUrl)
            : this()
        {
            this.FileUrl = FileUrl;
        }

        public abstract void Parse();
    }
}
