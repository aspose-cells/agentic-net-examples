using System;
using System.IO;
using Aspose.Cells;

class ExportShapesXml
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure XML save options (export all sheets)
            XmlSaveOptions xmlOptions = new XmlSaveOptions
            {
                SheetIndexes = null
            };

            string outputPath = "shapes_definitions.xml";

            // Save the workbook as XML containing shape definitions
            workbook.Save(outputPath, xmlOptions);

            Console.WriteLine($"All shape XML definitions have been exported to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}