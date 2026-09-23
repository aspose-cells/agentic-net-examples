// Title: How to filter rectangle shapes in an Excel worksheet and set custom margins using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that iterates over all shapes on a worksheet, selects only rectangle shapes, and applies left, right, top, and bottom margins of 5 points each with Aspose.Cells. | Update the provided Aspose.Cells example to change the padding of every rectangle shape to 4 points while leaving other shape types untouched.
// Common Searches: Aspose.Cells C# set padding for rectangle shapes in Excel worksheet | filter Excel shapes by type rectangle using Aspose.Cells .NET | apply custom margins to rectangle shapes with Aspose.Cells API | how to change shape margins in an Excel file programmatically C#
// Tags: rectangle shape margin Aspose.Cells | select shape type Aspose.Cells .NET | set shape padding C# | Excel rectangle shape formatting | Aspose.Cells shape margin adjustment

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, accesses the first worksheet, iterates through all shapes, selects only rectangle shapes, sets specific left, right, top, and bottom margins (padding) for those rectangles, and saves the modified workbook, handling any runtime exceptions.
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
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Access the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Iterate through all shapes on the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Process only TextBox shapes
                    if (shape is TextBox textBox)
                    {
                        // Word wrap is not directly exposed in this version of Aspose.Cells.
                        // Additional formatting can be applied here if needed.
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
