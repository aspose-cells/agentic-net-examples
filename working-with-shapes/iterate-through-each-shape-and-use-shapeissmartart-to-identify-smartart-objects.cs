// Title: Detect SmartArt Shapes in an Excel Workbook with Aspose.Cells for .NET (C#)
// Description: Load an existing .xlsx file, loop through each worksheet and its ShapeCollection, use the Shape.IsSmartArt property to spot SmartArt objects, output the worksheet name, shape ID and shape name, and optionally save the workbook. Ideal for reporting or further SmartArt manipulation.
// Keywords: Aspose.Cells SmartArt detection | C# Shape.IsSmartArt example | enumerate Excel shapes .NET | list SmartArt objects Aspose | identify SmartArt in workbook
// Common Searches: how to find SmartArt shapes using Aspose.Cells | Aspose.Cells Shape.IsSmartArt C# sample | list SmartArt objects in each Excel worksheet | enumerate shapes and detect SmartArt in .NET
// Developer Intent: Locate and list every SmartArt shape inside an Excel workbook using Aspose.Cells.
// Use Cases: Generate a report of SmartArt IDs and names per worksheet. | Collect SmartArt shape IDs for batch processing or conversion. | Validate that a workbook contains required SmartArt elements before publishing.
// AI Prompts: Create C# code with Aspose.Cells that iterates all worksheets and prints details of SmartArt shapes. | Show how to filter a ShapeCollection for SmartArt objects and store their IDs in a list. | Explain how to change properties (e.g., size, position) of identified SmartArt shapes using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Load an existing .xlsx file, loop through each worksheet and its ShapeCollection, use the Shape.IsSmartArt property to spot SmartArt objects, output the worksheet name, shape ID and shape name, and optionally save the workbook. Ideal for reporting or further SmartArt manipulation.
    public class IdentifySmartArtShapes
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets in the workbook
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    // Get the collection of shapes on the current worksheet
                    ShapeCollection shapes = worksheet.Shapes;

                    // Iterate through each shape in the collection
                    foreach (Shape shape in shapes)
                    {
                        // Determine if the shape is a SmartArt object
                        if (shape.IsSmartArt)
                        {
                            Console.WriteLine(
                                $"SmartArt found in worksheet \"{worksheet.Name}\" with shape ID {shape.Id} and name \"{shape.Name}\".");
                        }
                    }
                }

                // Save the workbook (optional, if any modifications were made)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display the error message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            IdentifySmartArtShapes.Run();
        }
    }
}
