// Title: How to retrieve and display absolute positions of all shapes on an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get shape top and left coordinates from worksheet | list all shapes with their dimensions in an Excel file using Aspose.Cells | retrieve shape positions and sizes for each shape in a workbook with Aspose.Cells | C# code to enumerate worksheet shapes and output their absolute locations | display shape size and location for every shape in an Excel sheet using Aspose.Cells
// Tags: Aspose.Cells shape coordinate extraction | enumerate worksheet shapes .NET | retrieve shape dimensions Aspose.Cells | Excel shape position analysis C# | list shape properties Aspose.Cells workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace Example
{
    // The example verifies the workbook file exists, loads it with Aspose.Cells, accesses the first worksheet, iterates through the worksheet's Shapes collection, and writes each shape's Name, Type, Top, Left, Height, and Width (in points) to the console, with error handling for robustness.
    public class ShapePositionAnalyzer
    {
        public static void DisplayShapePositions(string workbookPath)
        {
            // Verify that the file exists before attempting to load it
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                return;
            }

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(workbookPath);

                // Access the first worksheet (change index or name as needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Iterate through all shapes on the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Output shape details to the console
                    Console.WriteLine($"Shape Name : {shape.Name}");
                    Console.WriteLine($"  Type   : {shape.Type}");
                    Console.WriteLine($"  Top    : {shape.Top} pt");
                    Console.WriteLine($"  Left   : {shape.Left} pt");
                    Console.WriteLine($"  Height : {shape.Height} pt");
                    Console.WriteLine($"  Width  : {shape.Width} pt");
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Accept workbook path from command‑line arguments or use a default placeholder
            string workbookPath = args.Length > 0 ? args[0] : "sample.xlsx";

            ShapePositionAnalyzer.DisplayShapePositions(workbookPath);
        }
    }
}
