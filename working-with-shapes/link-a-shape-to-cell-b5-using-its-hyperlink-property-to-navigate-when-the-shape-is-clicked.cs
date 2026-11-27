// Title: Assign an internal hyperlink to an AutoShape that jumps to cell B5 using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set hyperlink on rectangle shape to navigate to cell B5 | Create an AutoShape with a link that jumps to a specific worksheet cell in Aspose.Cells | Link a shape to an internal cell reference using Aspose.Cells .NET API | Add a clickable shape that opens cell B5 in an Excel file with Aspose.Cells | Set Hyperlink.Address '#B5' on a shape in Aspose.Cells C# example
// Tags: Aspose.Cells shape hyperlink to internal cell | C# AutoShape navigation using Hyperlink.Address | set shape hyperlink address '#B5' Aspose.Cells | link rectangle auto shape to worksheet cell | Aspose.Cells click-to-jump shape example

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, adds a rectangle AutoShape to the first worksheet, assigns it display text, and configures its Hyperlink.Address to "#B5" with a ScreenTip, enabling the shape to navigate to cell B5 when clicked. The workbook is saved as ShapeLinkedToB5.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle auto shape to the worksheet.
            // Parameters: AutoShapeType, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = worksheet.Shapes.AddAutoShape(
                AutoShapeType.Rectangle,
                5,   // upper left row (zero‑based)
                5,   // upper left column (zero‑based)
                5,   // top offset in pixels
                5,   // left offset in pixels
                100, // height in pixels
                200  // width in pixels
            );

            // Set shape properties
            shape.Name = "NavigateToB5";
            shape.AlternativeText = "Click to navigate to cell B5";
            shape.Text = "Go to B5";

            // Configure hyperlink to cell B5.
            // For internal cell navigation, set the address to "#B5".
            Hyperlink hyperlink = shape.Hyperlink;
            hyperlink.Address = "#B5";
            hyperlink.ScreenTip = "Go to B5";

            // Save the workbook
            string outputPath = "ShapeLinkedToB5.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved to '{outputPath}'. The shape is linked to cell B5.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
