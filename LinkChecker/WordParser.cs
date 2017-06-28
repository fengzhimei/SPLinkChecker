using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace LinkChecker
{
    class WordParser : BaseParser
    {
        public WordParser()
            : base()
        { 
        }

        public override void Parse()
        {
            StringBuilder hyperlinkText = null;
            string hyperlinkRelationshipId = string.Empty;

            try
            {
                byte[] fileByteArray = Utility.GetFileByteArray(FileUrl);

                if (fileByteArray != null)
                {
                    using (MemoryStream fileStream = new MemoryStream(fileByteArray, false))
                    {
                        using (WordprocessingDocument doc = WordprocessingDocument.Open(fileStream, false))
                        {
                            Document mainDocument = doc.MainDocumentPart.Document;

                            // Iterate through the hyperlink elements in the
                            // main document part.
                            foreach (DocumentFormat.OpenXml.Wordprocessing.Hyperlink hyperlink in mainDocument.Descendants<DocumentFormat.OpenXml.Wordprocessing.Hyperlink>())
                            {
                                if (hyperlink.Id != null)
                                {
                                    hyperlinkText = new StringBuilder();

                                    // Get the text in the document that is associated
                                    // with the hyperlink. The text could be spread across
                                    // multiple text elements so process all the text
                                    // elements that are descendants of the hyperlink element.
                                    foreach (DocumentFormat.OpenXml.Wordprocessing.Text text in hyperlink.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                                        hyperlinkText.Append(text.InnerText);

                                    // The hyperlink element has an explicit relationship
                                    // with the actual hyperlink. Get the relationship id
                                    // via the hyperlink element's Id attribute.
                                    hyperlinkRelationshipId = hyperlink.Id.Value;

                                    // Get the hyperlink uri via the explicit relationship Id.
                                    HyperlinkRelationship hyperlinkRelationship = doc
                                        .MainDocumentPart.HyperlinkRelationships
                                        .Single(c => c.Id == hyperlinkRelationshipId);

                                    if (hyperlinkRelationship != null
                                        && hyperlinkRelationship.IsExternal)
                                    {
                                        if (hyperlinkRelationship.Uri.IsAbsoluteUri == false)
                                        {
                                            FileLink objLink = new FileLink()
                                            {
                                                ParentFileUrl = FileUrl,
                                                LinkText = hyperlinkText.ToString(),
                                                LinkAddress = hyperlinkRelationship.Uri.OriginalString,
                                                hasError = true
                                            };

                                            Results.Add(objLink);
                                        }
                                        else
                                        {
                                            if (hyperlinkRelationship.Uri.Scheme != Uri.UriSchemeMailto)
                                            {
                                                FileLink objLink = new FileLink()
                                                {
                                                    ParentFileUrl = FileUrl,
                                                    LinkText = hyperlinkText.ToString(),
                                                    LinkAddress = hyperlinkRelationship.Uri.AbsoluteUri
                                                };

                                                Results.Add(objLink);
                                            }
                                        }
                                    }
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
