using System;
using Aspose.Cells;

namespace AsposeCellsXmlMapImport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook (any supported Excel format)
            string inputPath = "source.xlsx";

            // Load the workbook using the standard load rule
            Workbook workbook = new Workbook(inputPath);

            // At this point the workbook is ready for further XML map import operations
            // Example: you can later call workbook.ImportXml(...) or work with XmlLoadOptions as needed

            Console.WriteLine("Workbook loaded successfully. Worksheets count: " + workbook.Worksheets.Count);
        }
    }
}