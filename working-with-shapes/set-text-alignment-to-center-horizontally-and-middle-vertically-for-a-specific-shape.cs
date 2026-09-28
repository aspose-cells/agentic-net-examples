// Title: Center text horizontally and vertically in a specific shape using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel workbook, locate a shape by index or name, and set its TextHorizontalAlignment and TextVerticalAlignment properties to Center with Aspose.Cells in C#. | Programmatically align the text inside a shape to the middle on both axes, then save the updated workbook using Aspose.Cells for .NET. | Retrieve the first shape on the first worksheet, apply centered alignment for both horizontal and vertical text, and write the changes to a new file with Aspose.Cells.
// Common Searches: aspnet aspose.cells set shape text alignment to center | c# aspose.cells align shape text horizontally and vertically | how to center text inside an Excel shape using Aspose.Cells library | aspose.cells change text alignment of a shape programmatically | example code for TextHorizontalAlignment and TextVerticalAlignment in C#
// Tags: Aspose.Cells shape alignment settings | Aspose.Cells shape text positioning | Excel shape formatting with Aspose | C# modify shape properties Aspose.Cells | Aspose.Cells text alignment API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The program loads an existing Excel file, accesses the first shape on the first worksheet, sets its TextHorizontalAlignment and TextVerticalAlignment to Center, and saves the modified workbook to a new file.
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

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one shape in the worksheet
                if (worksheet.Shapes.Count == 0)
                {
                    Console.WriteLine("No shapes found in the worksheet.");
                    return;
                }

                // Retrieve the first shape (or you could use a specific name)
                Shape shape = worksheet.Shapes[0];

                // Set horizontal and vertical text alignment to center
                shape.TextHorizontalAlignment = TextAlignmentType.Center;
                shape.TextVerticalAlignment = TextAlignmentType.Center;

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
