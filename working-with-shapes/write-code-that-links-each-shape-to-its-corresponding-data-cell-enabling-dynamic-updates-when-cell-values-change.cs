// Title: How to link Excel shapes to specific cells for dynamic text updates using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# link shape text to a cell value dynamically | set Excel shape formula programmatically with Aspose.Cells | update shape text when worksheet cell changes using Aspose.Cells .NET | map multiple shape names to cell addresses in C# Aspose.Cells example | configure shape placement to move and size with cells Aspose.Cells
// Tags: shape text formula binding Aspose.Cells | link shape to cell address C# | dynamic shape text update Excel Aspose.Cells | shape placement MoveAndSize Aspose.Cells | dictionary mapping shapes to cells Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeLinkExample
{
    // The C# example loads an existing workbook, defines a dictionary that maps shape names to cell addresses, iterates through all shapes on the first worksheet, and for each matching shape sets its Text property to a formula like "=A1" so the shape displays the cell's value dynamically. It also sets the shape's Placement to MoveAndSize and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Mapping of shape names to target cell addresses
                var shapeCellMap = new Dictionary<string, string>()
                {
                    { "Shape1", "A1" },
                    { "Shape2", "B2" },
                    { "Shape3", "C3" }
                    // Add more mappings as required
                };

                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // If the shape name is in the dictionary, link its text to the cell
                    if (shapeCellMap.TryGetValue(shape.Name, out string cellAddress))
                    {
                        // The leading '=' makes the shape display the cell's value dynamically
                        shape.Text = "=" + cellAddress;

                        // Set the shape to move and size with cells
                        shape.Placement = PlacementType.MoveAndSize;
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
