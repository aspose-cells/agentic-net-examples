// Title: Batch update the LinkedCell property of shapes across all worksheets with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# batch update LinkedCell of shapes in every worksheet | Programmatically change linked cell address of Excel shapes using Aspose.Cells | Iterate through all worksheets and shapes to set new LinkedCell values in .NET | Use a dictionary to map shape names to new linked cells with Aspose.Cells | Save an Excel workbook after bulk updating shape LinkedCell references in C#
// Tags: batch update shape LinkedCell Aspose.Cells | iterate worksheets and shapes C# | dictionary driven linked cell mapping | save workbook after shape modifications Aspose.Cells | linked shape cell reference update .NET | process only shapes with non-empty LinkedCell

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace LinkedShapeUpdater
{
    // The example loads an existing Excel workbook, walks through each worksheet and its shapes, and replaces the LinkedCell of any shape whose name is present in a provided dictionary with a new cell address, then saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the existing workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Define new linked cell addresses for shapes (key: shape name, value: new linked cell address)
            var newLinkedCells = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "LinkedShape1", "DataSheet!$B$2" },
                { "LinkedShape2", "DataSheet!$C$5" },
                { "LinkedShape3", "DataSheet!$D$10" }
                // Add more shape mappings as needed
            };

            try
            {
                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes in the current worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Process only shapes that are linked to a cell (LinkedCell property is non‑empty)
                        if (!string.IsNullOrEmpty(shape.LinkedCell))
                        {
                            // Update linked cell if the shape name exists in the dictionary
                            if (newLinkedCells.TryGetValue(shape.Name, out string newAddress))
                            {
                                shape.LinkedCell = newAddress;
                            }
                        }
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while updating shapes or saving the workbook: {ex.Message}");
            }
        }
    }
}
