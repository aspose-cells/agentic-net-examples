// Title: Iterate through all worksheet shapes and set a uniform font size for shape text using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to loop through worksheet shapes and change text font size | set same font size for all shape text effects in an Excel file with Aspose.Cells | Aspose.Cells .NET iterate shapes and apply text formatting to each shape | update TextEffect font size for every shape in a workbook using C# | batch modify shape text properties in Excel using Aspose.Cells API
// Tags: worksheet shape iteration Aspose.Cells | TextEffect font size update | batch modify shape text properties | Aspose.Cells uniform shape font size | C# Excel shape formatting loop

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads a workbook, accesses the first worksheet, iterates over each Shape, checks for a TextEffect format, sets the TextEffect FontSize to 12 points, and saves the modified workbook.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (or specify another index/name as needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Iterate over all shapes in the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Ensure the shape has a TextEffect format before modifying
                    if (shape.TextEffect != null)
                    {
                        // Adjust a valid property, e.g., FontSize, which is supported by TextEffectFormat
                        shape.TextEffect.FontSize = 12; // Set font size to 12 points
                    }
                }

                // Save the modified workbook
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
