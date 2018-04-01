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
        static void Main(string[] args)
        {
            XElement toc = WDRetrieveTOC(@"C:\Users\jeremy.leff\Desktop\Contents.docx");

            //IEnumerable<XElement> links =
            //    from el in toc.Descendants()
            //    where el.Name.Namespace == "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            //    select el;
            Console.WriteLine("All Elements:\n");
            foreach (XElement el in toc.Descendants())
            {
                
                Console.WriteLine(el.Value);
                Console.WriteLine("\n");

            }
                

            //foreach (XElement el in links)
            //    Console.WriteLine(el);

            //Console.WriteLine(toc.Value.ToString());
        }

        public static XElement WDRetrieveTOC(string fileName)
        {
            XElement TOC = null;

            using (var document = WordprocessingDocument.Open(fileName, false))
            {
                var docPart = document.MainDocumentPart;
                var doc = docPart.Document;

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
            }
            return TOC;
        }
    }
}
