using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml.Linq;

namespace DocParseTest
{
    class Program
    {
        private static XNamespace w = XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main");

        static void Main(string[] args)
        {
            Document doc = OpenDocument(@"C:\Users\jeremy.leff\Desktop\Contents.docx");
            XElement toc = WDRetrieveTOC(doc);
            var contents = ParseTOC(toc);
            bool startContentSection = false;

            foreach(var c in contents)
                Console.WriteLine(c.Value + ": " + c.Key);

            var xdoc = XElement.Parse(doc.OuterXml);

            var paragraphs =
                from para in xdoc.Elements(w + "body")
                    .FirstOrDefault()
                    .Elements(w + "p")
                select para;

            foreach (var p in paragraphs)
            {
                var bookmarks =
                from bm in p.Elements(w + "bookmarkStart")
                select bm;

                foreach (var bm in bookmarks)
                {
                    
                    if (contents.ContainsKey(bm.Attribute(w + "name").Value))
                    {
                        startContentSection = true;
                    }
                }
            }

            

            

        }

        public static Dictionary<string, string> ParseTOC(XElement toc)
        {
            var contents = new Dictionary<string, string>();
            

            IEnumerable<XElement> links =
                from el in toc.Descendants()
                where el.Name.Namespace == "http://schemas.openxmlformats.org/wordprocessingml/2006/main" &&
                el.Name.LocalName == "hyperlink"
                select el;

            foreach (XElement l in links)
            {
                XElement e  = (XElement)l.FirstNode;
                XElement f  = (XElement)e.FirstNode.NextNode;
                var val     = f.Value;

                contents.Add(l.Attribute(w + "anchor").Value, val);
            }

            return contents;
        }

        public static XElement WDRetrieveTOC(Document doc)
        {
            XElement TOC = null;

            OpenXmlElement block = doc.Descendants<DocPartGallery>().
                Where(b => b.Val.HasValue &&
                (b.Val.Value == "Table of Contents")).FirstOrDefault();

            if (block != null)
            {
                // Back up to the enclosing SdtBlock and return that XML.
                while ((block != null) && (!(block is SdtBlock)))
                {
                    block = block.Parent;
                }
                TOC = XElement.Parse(block.OuterXml);
            }
            
            return TOC;
        }

        public static Document OpenDocument(string fileName)
        {
            using (var document = WordprocessingDocument.Open(fileName, false))
            {
                var docPart = document.MainDocumentPart;
                var doc = docPart.Document;

                return doc;
            }
        }
    }
}
