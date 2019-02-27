using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
namespace Open_XML_Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 2)
            {
                string filepath = args[0];
                WriteToDocument(filepath);

            }
            else if (args.Length == 1)
            {
                string filepath = args[0];
                Console.WriteLine("Reading from " + filepath);
                Console.WriteLine(ReadFromDocument(filepath));
                Console.ReadLine();
            }
        }


        static void WriteToDocument(string filepath)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {
                var mainDocumentPart = wordDocument.AddMainDocumentPart();
                mainDocumentPart.Document = new Document(new Body(new Paragraph(new Run(new Text("Hello Open XML World!")))));
            }
        }

        static string ReadFromDocument(string path)
        {
            //string path = "C:\\Users\\Ferid\\Desktop\\1 abstract.docx";
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(path, false))
            {
                return wordDocument.MainDocumentPart.Document.InnerText;
            }
        }
    }
}
