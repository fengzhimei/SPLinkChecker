using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace LinkChecker
{
    class PowerPointParser : BaseParser
    {
        public PowerPointParser()
            : base()
        { 
        }

        public override void Parse()
        {
            string hyperlinkText = string.Empty;
            string hyperlinkRelationshipId;

            try
            {
                byte[] fileByteArray = Utility.GetFileByteArray(FileUrl);

                if (fileByteArray != null)
                {
                    using (MemoryStream fileStream = new MemoryStream(fileByteArray, false))
                    {
                        using (PresentationDocument document = PresentationDocument.Open(fileStream, false))
                        {
                            // Iterate through all the slide parts in the presentation part.
                            foreach (SlidePart slidePart in document.PresentationPart.SlideParts)
                            {
                                IEnumerable<DocumentFormat.OpenXml.Drawing.HyperlinkType> links = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.HyperlinkType>();

                                // Iterate through all the links in the slide part.
                                foreach (DocumentFormat.OpenXml.Drawing.HyperlinkType hyperlink in links)
                                {
                                    if (hyperlink.Id != null)
                                    {
                                        hyperlinkText = Utility.GetPPTHyperlinkText(hyperlink);
                                        hyperlinkRelationshipId = hyperlink.Id.Value;
                                        HyperlinkRelationship hyperlinkRelationship = slidePart
                                            .HyperlinkRelationships
                                            .Single(c => c.Id == hyperlinkRelationshipId);
                                        if (hyperlinkRelationship != null
                                            && hyperlinkRelationship.IsExternal)
                                        {
                                            if (hyperlinkRelationship.Uri.IsAbsoluteUri == false)
                                            {
                                                FileLink objLink = new FileLink()
                                                {
                                                    ParentFileUrl = FileUrl,
                                                    LinkText = hyperlinkText,
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
                                                        LinkText = hyperlinkText,
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
