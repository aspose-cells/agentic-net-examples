// Title: How to align a shape with a header cell in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# align shape with header cell in worksheet | set shape UpperLeftRow UpperLeftColumn to cell A1 using Aspose.Cells | position Excel shape over a specific cell programmatically with Aspose.Cells | C# move rectangle shape to top-left corner of sheet in Aspose.Cells | adjust shape absolute coordinates in an Excel file using Aspose.Cells for .NET
// Tags: shape back ordering Aspose.Cells | position shape using UpperLeftRow UpperLeftColumn Aspose.Cells | shape layering order Excel C# | align shape with header cell Aspose.Cells | adjust shape absolute coordinates Excel .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, adds a rectangle shape if none exists, retrieves the first shape, sets its UpperLeftRow and UpperLeftColumn to 0 to position it over cell A1 (the header cell), and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists; create a sample if missing
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found. Creating a sample workbook.");
                    Workbook sampleWb = new Workbook();
                    Worksheet sampleWs = sampleWb.Worksheets[0];
                    // Add a simple rectangle shape as a placeholder
                    sampleWs.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 50);
                    sampleWb.Save(inputPath);
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure at least one shape exists; add one if necessary
                if (sheet.Shapes.Count == 0)
                {
                    Console.WriteLine("No shapes found. Adding a default rectangle shape.");
                    sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 0, 0, 0, 0, 100, 50);
                }

                // Retrieve the first shape
                Shape shape = sheet.Shapes[0];

                try
                {
                    // Align the shape with cell A1 (row 0, column 0)
                    shape.UpperLeftRow = 0;
                    shape.UpperLeftColumn = 0;
                    // Offsets default to 0; explicit setting is optional
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error adjusting shape position: {ex.Message}");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
