// Title: Copy a shape's absolute row/column and offset coordinates to another shape in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that reads the UpperLeftRow, UpperLeftColumn, Top, and Left of a source Shape and applies those values to a target Shape with Aspose.Cells. | Show how to guarantee two rectangle shapes exist in a worksheet, then overlay the second shape exactly on the first by copying Width and Height properties using Aspose.Cells. | Provide a complete example that loads (or creates) a workbook, aligns a target shape to a source shape's absolute position, and saves the workbook as an .xlsx file.
// Common Searches: how to align one shape to another in Aspose.Cells C# | copy shape coordinates from one rectangle to another using Aspose.Cells .NET | retrieve absolute row and column of a shape in Excel with Aspose.Cells | overlay two shapes at the same position in an Excel file using Aspose.Cells | Aspose.Cells set shape Top and Left properties from another shape
// Tags: Aspose.Cells copy shape position | Aspose.Cells set shape coordinates | Aspose.Cells overlay shapes Excel | Aspose.Cells shape size synchronization | Aspose.Cells workbook shape alignment

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeAlignmentDemo
{
    // // Loads or creates a workbook, ensures two rectangle shapes exist, copies the source shape's UpperLeftRow, UpperLeftColumn, Top, Left, Width, and Height to the target shape to align them, and saves the result as an .xlsx file.
    class ShapeAlignment
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new workbook
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // creates a default workbook with one worksheet
                }

                Worksheet sheet = workbook.Worksheets[0];
                ShapeCollection shapes = sheet.Shapes;

                // Ensure at least two shapes are present
                if (shapes.Count < 2)
                {
                    // Add a source rectangle shape
                    shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 50);
                    // Add a target rectangle shape
                    shapes.AddShape(MsoDrawingType.Rectangle, 5, 5, 0, 0, 100, 50);
                }

                // Retrieve the first two shapes
                Shape source = shapes[0];
                Shape target = shapes[1];

                // Align target shape to source shape's absolute position
                target.UpperLeftRow = source.UpperLeftRow;
                target.UpperLeftColumn = source.UpperLeftColumn;
                target.Top = source.Top;   // vertical offset in points
                target.Left = source.Left; // horizontal offset in points

                // Optionally copy size for exact overlay
                target.Width = source.Width;
                target.Height = source.Height;

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
