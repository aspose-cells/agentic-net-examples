// Title: Read a shape's width and height from an Excel worksheet and write the values to cells using Aspose.Cells for .NET
// Common Searches: how to retrieve shape width and height in an Excel file with Aspose.Cells C# | write shape dimensions to specific cells in a workbook using Aspose.Cells for .NET | Aspose.Cells get shape size and store in A1 and B1 cells C# example | C# code to read shape bounding box dimensions from Excel and save to a new file | extract shape dimensions from the first worksheet shape using Aspose.Cells
// Tags: Aspose.Cells shape width extraction | Aspose.Cells shape height extraction | write shape size to worksheet cells C# | read shape bounding box dimensions Aspose.Cells | export shape dimensions to Excel cells .NET | process first worksheet shape Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, checks for at least one shape on the first worksheet, reads the shape's Width and Height properties, writes the width to cell A1 and the height to cell B1, and saves the modified workbook to a new file.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify input file existence
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one shape
                if (sheet.Shapes.Count == 0)
                {
                    Console.WriteLine("No shapes found in the worksheet.");
                    return;
                }

                // Get the first shape
                Shape shape = sheet.Shapes[0];

                // Use shape's Width and Height as the bounding dimensions
                float width = shape.Width;
                float height = shape.Height;

                // Write dimensions to cells A1 (width) and B1 (height)
                sheet.Cells["A1"].PutValue(width);
                sheet.Cells["B1"].PutValue(height);

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Dimensions saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
