// Title: How to bring a SmartArt shape to the front by setting its Z‑order in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, iterates through worksheet.Shapes, detects SmartArt objects, and sets their ZOrderPosition to 0 to place them on the top layer. | Provide a C# example showing how to modify the Z‑order of a specific SmartArt shape so it appears above all other shapes in an Excel file using Aspose.Cells.
// Common Searches: Aspose.Cells C# set SmartArt ZOrderPosition to front | How to change shape layering in Excel with Aspose.Cells .NET | Bring SmartArt to top layer programmatically using Aspose.Cells | C# Aspose.Cells move shape to front of other objects | Set Z‑order of a SmartArt shape in a workbook with Aspose.Cells
// Tags: smartart z-order Aspose.Cells | set shape ZOrderPosition .NET | bring shape to front Excel | modify shape layering Aspose.Cells | worksheet shapes ordering C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, loops through all shapes on the first worksheet, identifies SmartArt shapes, sets their ZOrderPosition to 0 to bring them to the front, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (or any specific worksheet)
                Worksheet worksheet = workbook.Worksheets[0];

                // Iterate through all shapes in the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Check if the shape is a SmartArt
                    if (shape.IsSmartArt)
                    {
                        // Bring the SmartArt shape to the front by setting its Z‑order position to the lowest value (0)
                        shape.ZOrderPosition = 0;
                        // If you need to target a specific SmartArt, add additional checks here (e.g., by name) and break.
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
