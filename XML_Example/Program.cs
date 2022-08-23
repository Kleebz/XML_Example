using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace XML_Example
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string xmlPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"SampleData\books.xml");

            Examples examples = new Examples();

            var xml1 = examples.DeserializeXML<catalog>(xmlPath);
            var xml2 = examples.DeserializeXML(xmlPath);

            string xml3 = examples.SerializeXML(xml1);

            examples.LinqQuery(xml2);

            Console.WriteLine("\n\n\n\n");
            Console.WriteLine("Where would you like to save a Word Doc?");
            string savePath = Console.ReadLine();
            examples.WordDocExample(xmlPath, savePath);
        }
    }
}
