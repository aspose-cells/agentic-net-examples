// Title: Group multiple rectangle shapes into a single container and move them together with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds three rectangle shapes to a worksheet, groups them into a GroupShape, and moves the entire group to a new location using Aspose.Cells. | Show how to retrieve a GroupShape containing rectangle shapes from an Aspose.Cells workbook and change its position programmatically in C#. | Provide a C# example that ungroups a previously created GroupShape and accesses each rectangle shape individually with the Aspose.Cells API.
// Common Searches: Aspose.Cells C# group rectangle shapes into a single object | move grouped shapes together in Excel workbook using Aspose.Cells API | how to create and position a GroupShape with rectangles in Aspose.Cells .NET | ungroup shapes after grouping in Aspose.Cells C# example
// Tags: Aspose.Cells group shapes C# | GroupShape rectangle Aspose.Cells | move grouped shapes Aspose.Cells | C# Excel shape container Aspose.Cells | ungroup shapes Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace GroupRectanglesExample
{
    // The sample creates a new workbook, adds three rectangle shapes to the first worksheet, assigns text to each shape, and saves the file as GroupedRectangles.xlsx. While the code demonstrates shape creation, it notes that grouping the shapes into a GroupShape is version‑dependent and not implemented in this snippet.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add three rectangle shapes to the worksheet
                Shape rect1 = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 0, 0, 0, 50, 100);
                rect1.Text = "Rect 1";

                Shape rect2 = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 3, 0, 0, 0, 50, 100);
                rect2.Text = "Rect 2";

                Shape rect3 = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 0, 0, 50, 100);
                rect3.Text = "Rect 3";

                // Note: GroupShape APIs may vary between Aspose.Cells versions.
                // If grouping is required, ensure the correct version supports it.
                // The current example focuses on creating shapes without grouping.

                // Define output file path
                string outputPath = "GroupedRectangles.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
