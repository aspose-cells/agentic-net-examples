// Title: How to apply a 50% transparent reflection to a specific shape in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set shape reflection transparency to 50 percent | Apply reflection effect with 0.5 transparency to a named shape in Excel using Aspose.Cells | Aspose.Cells shape.Reflection.Transparency example for .NET | How to add a semi‑transparent reflection to an Excel shape programmatically | Modify shape reflection properties in a workbook with Aspose.Cells C#
// Tags: Aspose.Cells shape reflection transparency | C# shape reflection API Aspose.Cells | Excel shape visual effect Aspose.Cells | apply semi-transparent reflection .NET | update shape reflection settings C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, retrieves the shape named "MyShape" from the first worksheet, sets its Reflection.Transparency to 0.5 (50% transparent), and saves the workbook, handling missing files and ensuring the output directory exists.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (or any specific worksheet)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the shape by name
                Shape shape = worksheet.Shapes["MyShape"];
                if (shape == null)
                {
                    Console.WriteLine("Shape 'MyShape' not found in the worksheet.");
                    return;
                }

                // Apply a reflection effect with 50% transparency
                // The reflection type defaults to a standard reflection; only transparency is set explicitly
                shape.Reflection.Transparency = 0.5; // 0.5 = 50% transparent

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the updated shape
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
