using HtmlAgilityPack;
using LumiSoft.Net.Mail;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace LinkChecker
{
    class HtmlParser : BaseParser
    {
        public HtmlParser()
            : base()
        {
            
        }

        public override void Parse()
        {
            string linkText = string.Empty;
            string linkURL = string.Empty;

            try
            {
                byte[] fileByteArray = Utility.GetFileByteArray(FileUrl);

                if (fileByteArray != null)
                {
                    using (MemoryStream fileStream = new MemoryStream(fileByteArray, false))
                    {
                        Mail_Message mime = Mail_Message.ParseFromStream(fileStream);
                        HtmlDocument doc = new HtmlDocument();
                        string html = string.IsNullOrEmpty(mime.BodyHtmlText) ? mime.BodyText : mime.BodyHtmlText;
                        if (!string.IsNullOrEmpty(html))
                        {
                            doc.LoadHtml(html);
                            foreach (HtmlNode link in doc.DocumentNode.SelectNodesOrEmpty("//a[@href]"))
                            {
                                if (link.Attributes["href"].Value.StartsWith("#", StringComparison.InvariantCultureIgnoreCase) == false
                                    && link.Attributes["href"].Value.StartsWith("javascript:", StringComparison.InvariantCultureIgnoreCase) == false
                                    && link.Attributes["href"].Value.StartsWith("mailto:", StringComparison.InvariantCultureIgnoreCase) == false)
                                {
                                    linkURL = link.Attributes["href"].Value;
                                    if (link.FirstChild == link.LastChild)
                                    {
                                        linkText = link.InnerText;
                                    }
                                    else
                                    {
                                        linkText = link.LastChild.InnerText;
                                    }

                                    linkText = new string(linkText.ToCharArray()).Replace("\r\n", " ");

                                    FileLink objLink = new FileLink()
                                    {
                                        ParentFileUrl = FileUrl,
                                        LinkText = linkText,
                                        LinkAddress = linkURL
                                    };

                                    Results.Add(objLink);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                FileLink objLink = new FileLink()
                {
                    ParentFileUrl = FileUrl,
                    LinkText = "Error occurred when parsing this file.",
                    LinkAddress = ex.Message,
                    hasError = true
                };

                Results.Add(objLink);
            }
        }
    }
}
