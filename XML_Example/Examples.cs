using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.IO;
using System.Xml.Linq;
using System.Xml;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace XML_Example
{
    public class Examples
    {
        //Deserialize XML to C# class
        public catalog DeserializeXML(string filePath)//Example sourced from: https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/concepts/serialization/how-to-read-object-data-from-an-xml-file
        {
            StreamReader file = new StreamReader(filePath);
            XmlSerializer reader = new XmlSerializer(typeof(catalog));
            catalog catalog = (catalog)reader.Deserialize(file);

            file.Dispose();

            return catalog;
        }

        //Generic way of the above example. Just wanted to include one generic example, but this can be applied to almost any method.
        public T DeserializeXML<T>(string filePath)//Example sourced from: https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/concepts/serialization/how-to-read-object-data-from-an-xml-file
        {
            T xmlClass;

            using (StreamReader file = new StreamReader(filePath))//using block automatically callse .Dispose()
            {
                XmlSerializer reader = new XmlSerializer(typeof(T));
                xmlClass = (T)reader.Deserialize(file);
            }

            return xmlClass;
        }

        public string SerializeXML(catalog catalog)//source: https://stackoverflow.com/questions/4123590/serialize-an-object-to-xml
        {
            XmlSerializer serializer = new XmlSerializer(typeof(catalog));
            var xml = string.Empty;

            using (var sWriter = new StringWriter())
            {
                using (XmlWriter writer = XmlWriter.Create(sWriter))
                {
                    serializer.Serialize(writer, catalog);
                    xml = sWriter.ToString(); // Your XML
                }
            }

            return xml;
        }

        public void LinqQuery(catalog catalog)
        {
            var books1 = catalog.book.ToList();
            var books2 = catalog.book.Where(x => x.price > (decimal)6.00).ToList();
            var bookTitles = catalog.book.Select(x => x.title).ToList();
            





            var anonymousObject1 = catalog.book.Select(x => new { x.title, x.author }).ToList();         
            var anonymousObject2 = catalog.book
                .OrderByDescending(x => x.price)
                .Select(x => new { Title = x.title, Author = x.author, Price = x.price })
                .ToList();

            foreach (var book in anonymousObject2)
            {
                Console.WriteLine($"${book.Price}    {book.Title}, By: {book.Author}");
            }






            Console.WriteLine("\n\n\n\n");

            anonymousObject2.Reverse();
            anonymousObject2.ForEach(book => {
                Console.WriteLine($"${book.Price}    {book.Title}, By: {book.Author}");
            });






            Console.WriteLine("\n\n\n\n");

            var books3 = catalog.book
                .GroupBy(x => x.author)
                .Select(x => new { Author = x.Key, Count = x.Count() })
                .ToList();

            books3.ForEach(x => {
                Console.WriteLine($"{x.Author} has written {x.Count} books.");
            });
        }

        public void WordDocExample(string filePath, string savePath)//Source: https://qawithexperts.com/article/c-sharp/c-create-or-generate-word-document-using-docx/366
        {                                                           //the source also has details on how to insert an image or a table
                                                                    //Their github repo has more advanced examples: https://github.com/xceedsoftware/docx
            var catalog = DeserializeXML(filePath);

            var doc = DocX.Create(savePath); //create docx word document

            doc.AddHeaders(); //add header (optional)
            doc.AddFooters(); //add footer in this document (optional code)

            // Force the first page to have a different Header and Footer.
            doc.DifferentFirstPage = true;

            doc.Headers.First.InsertParagraph("This is the ").Append("first").Bold().Append(" page header");// Insert a Paragraph into the first Header.
            doc.Footers.First.InsertParagraph("Page ").AppendPageNumber(PageNumberFormat.normal).Append(" of ").AppendPageCount(PageNumberFormat.normal); // add footer with page number


            var books = catalog.book.OrderByDescending(x => x.price).ToList();

            books.ForEach(book => {
                doc.InsertParagraph($"{book.title} | {book.author}\n");
                doc.InsertParagraph(book.description); // inserts a new paragraph with text
            });

            doc.Save(); // save changes to file
        }
    }
}
