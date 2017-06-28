using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LinkChecker
{
    class FileLink
    {
        public string ParentFileUrl
        {
            get;
            set;
        }

        public string LinkText
        {
            get;
            set;
        }

        public string LinkAddress
        {
            get;
            set;
        }

        public bool hasError
        {
            get;
            set;
        }
    }
}
