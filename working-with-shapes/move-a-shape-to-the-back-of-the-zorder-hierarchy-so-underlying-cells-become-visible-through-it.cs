// Title: How to send a named shape to the back of the Z‑order in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, finds a shape by its name, and sets its ZOrderPosition to 0 to place it behind all other objects. | Generate a .NET snippet that moves a specific shape to the back of the Z‑order so the cells underneath become visible, using the Shape.ZOrderPosition property. | Create a C# example that checks for a shape named 'MyShape' in the first worksheet and changes its Z‑order to the lowest level before saving the file.
// Common Searches: aspnet aspocells move shape to back of z-order c# | c# Aspose.Cells set shape behind cells in worksheet | how to change shape layering order in Excel file using Aspose.Cells .NET | retrieve shape by name and send to back Aspose.Cells example | make cells visible through shape Aspose.Cells C#
// Tags: Aspose.Cells shape Z‑order control | C# send Excel shape to back with Aspose.Cells | worksheet object layering using Aspose.Cells | expose cells hidden by shape in .xlsx | adjust shape order in Aspose.Cells workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads 'input.xlsx', accesses the first worksheet, retrieves the shape named 'MyShape', sets its ZOrderPosition to 0 to move it to the back of the Z‑order so underlying cells become visible, and saves the result as 'output.xlsx'.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";
            string shapeName = "MyShape";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Retrieve the shape by name
                Shape shape = sheet.Shapes[shapeName];

                if (shape == null)
                {
                    Console.WriteLine($"Shape '{shapeName}' not found in the worksheet.");
                }
                else
                {
                    // Send the shape to the back of the Z‑order by setting its position to 0
                    shape.ZOrderPosition = 0;
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
