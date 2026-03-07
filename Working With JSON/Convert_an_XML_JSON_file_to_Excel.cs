using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input file can be either a JSON file or an XML file
            string inputPath = "input.json";   // change to your source file path
            string outputPath = "output.xlsx"; // desired Excel output path

            // Determine file type by extension
            string extension = Path.GetExtension(inputPath).ToLowerInvariant();

            Workbook workbook;

            if (extension == ".json")
            {
                // Load JSON file into a workbook
                // JsonLoadOptions can be customized if needed (e.g., MultipleWorksheets)
                JsonLoadOptions jsonOptions = new JsonLoadOptions();
                // Example: jsonOptions.MultipleWorksheets = true;
                workbook = new Workbook(inputPath, jsonOptions);
            }
            else if (extension == ".xml")
            {
                // Load XML file into a workbook
                // XmlLoadOptions can be customized (e.g., IsXmlMap, ConvertNumericOrDate)
                XmlLoadOptions xmlOptions = new XmlLoadOptions();
                // Example: xmlOptions.IsXmlMap = true;
                workbook = new Workbook(inputPath, xmlOptions);
            }
            else
            {
                throw new NotSupportedException("Only .json or .xml input files are supported.");
            }

            // Save the loaded workbook as an Excel file
            workbook.Save(outputPath);

            Console.WriteLine($"Conversion completed. Excel file saved to: {outputPath}");
        }
    }
}