// Title: How to enumerate connector (line) shapes in an Excel worksheet and log each shape’s name and type with Aspose.Cells for .NET
// AI Prompts: Generate C# code that walks through the Shapes collection of a worksheet, identifies line shapes as connectors, and writes their Name and MsoDrawingType to the console. | Create a reusable method in Aspose.Cells that returns a list of connector shape names together with their drawing type for a given workbook.
// Common Searches: Aspose.Cells C# enumerate line shapes in an Excel worksheet | how to detect connector shapes using Aspose.Cells .NET | list shape names and types from Excel file with Aspose.Cells C#
// Tags: Aspose.Cells enumerate connector shapes | C# retrieve line shape properties Excel | log shape name and drawing type Aspose.Cells | detect connector shapes worksheet .NET | iterate worksheet shapes collection C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads an Excel workbook, scans the first worksheet’s Shapes collection, treats line shapes as connectors, and writes each connector’s Name and MsoDrawingType to the console; the workbook is then saved.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all shapes on the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                try
                {
                    // Identify possible connector shapes.
                    // Aspose.Cells does not expose explicit connector properties; treat line shapes as connectors.
                    bool isConnector = (int)shape.Type == (int)MsoDrawingType.Line;

                    if (isConnector)
                    {
                        Console.WriteLine($"Connector Shape: {shape.Name}");
                        Console.WriteLine($"  Shape Type: {shape.Type}");
                    }
                }
                catch (Exception exShape)
                {
                    Console.WriteLine($"Error processing shape '{shape?.Name}': {exShape.Message}");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (optional if modifications were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
