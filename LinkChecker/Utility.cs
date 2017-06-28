using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using HtmlAgilityPack;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace LinkChecker
{
    internal static class HtmlAgilityPackExtension
    {
        internal static IEnumerable<HtmlNode> SelectNodesOrEmpty(this HtmlNode htmlNode, String xpath)
        {
            return htmlNode.SelectNodes(xpath) ?? Enumerable.Empty<HtmlNode>();
        }
    }

    class Utility
    {
        internal static byte[] GetFileByteArray(string fileUrl)
        {
            Uri uri = null;

            if (Uri.TryCreate(fileUrl, UriKind.Absolute, out uri) && uri != null)
            {
                try
                {
                    HttpWebRequest hwReq = WebRequest.Create(uri) as HttpWebRequest;
                    hwReq.UseDefaultCredentials = true;

                    using (HttpWebResponse hwRes = hwReq.GetResponse() as HttpWebResponse)
                    {
                        if (hwRes != null && hwRes.StatusCode == HttpStatusCode.OK)
                        {
                            Stream hwResStream = hwRes.GetResponseStream();

                            int count = 0;
                            byte[] buffer = new byte[1024];

                            using (MemoryStream ms = new MemoryStream())
                            {
                                do
                                {
                                    count = hwResStream.Read(buffer, 0, buffer.Length);
                                    ms.Write(buffer, 0, count);
                                } while (hwResStream.CanRead && count > 0);
                                
                                return ms.ToArray();
                            }
                        }
                    }
                }
                catch { }
            }
            return null;
        }

        /// <summary>
        /// Travel back to parent container to get hyperlink text
        /// </summary>
        /// <param name="hyperlink"></param>
        /// <returns></returns>
        internal static string GetPPTHyperlinkText(HyperlinkType hyperlink)
        {
            if (string.IsNullOrEmpty(hyperlink.InnerText))
            {
                OpenXmlElement parent = hyperlink.Parent;
                while (string.IsNullOrEmpty(parent.InnerText))
                {
                    parent = parent.Parent;
                }
                return parent.InnerText;
            }
            return hyperlink.InnerText;
        }

        /// <summary>
        /// Force to use NTLM authentication
        /// remove all other authentication, keep "NTLM" only.
        /// AuthenticationType: Negotiate, Kerberos, NTLM, Digest, Basic
        /// </summary>
        internal static void ForceToUseNTLM()
        {
            IEnumerator registeredModules = AuthenticationManager.RegisteredModules;
            List<IAuthenticationModule> listRemove = new List<IAuthenticationModule>();
            while (registeredModules.MoveNext())
            {
                IAuthenticationModule currentAuthenticationModule = (IAuthenticationModule)registeredModules.Current;
                if (string.Compare(currentAuthenticationModule.AuthenticationType, "NTLM", true) != 0)
                {
                    listRemove.Add(currentAuthenticationModule);
                }
            }

            foreach (IAuthenticationModule m in listRemove)
            {
                AuthenticationManager.Unregister(m);
            }
        }
    }
}
