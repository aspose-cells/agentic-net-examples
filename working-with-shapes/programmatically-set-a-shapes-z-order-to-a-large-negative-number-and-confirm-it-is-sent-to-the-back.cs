// Title: Send a rectangle shape to the back of the Z‑order stack in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, adds a rectangle shape to the first worksheet, assigns its ZOrderPosition property a value of 0 to place it behind all other objects, and saves the workbook. | Write a C# snippet that retrieves an existing shape in an Aspose.Cells worksheet, sets its ZOrderPosition to a negative or zero value to move it to the back of the stacking order, and persists the changes.
// Common Searches: Aspose.Cells C# set shape ZOrderPosition to back of worksheet | move Excel shape behind other objects using Aspose.Cells .NET | change stacking order of shapes in Aspose.Cells workbook | C# Aspose.Cells send rectangle shape to back of Z-order | how to adjust shape Z-order in Excel file with Aspose.Cells API
// Tags: Aspose.Cells shape ZOrderPosition adjustment | C# set shape Z-order back Aspose.Cells | Excel shape stacking order manipulation .NET | rectangle shape Z-order Aspose.Cells example | programmatic shape ordering in Excel workbook

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangle shape to the first worksheet, sets its ZOrderPosition to 0 (the lowest value) so the shape is placed behind all other objects, saves the workbook as ZOrderDemo.xlsx, and writes status messages to the console.
class ShapeZOrderDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define rectangle parameters
            int upperLeftRow = 2;
            int upperLeftColumn = 2;
            int top = 5;
            int left = 5;
            int height = 100;
            int width = 200;

            // Add a rectangle shape to the worksheet
            Shape rectangle = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                upperLeftRow,
                upperLeftColumn,
                top,
                left,
                height,
                width);

            // Send the shape to the back of the Z‑order stack by setting its ZOrderPosition to the lowest value
            rectangle.ZOrderPosition = 0;

            Console.WriteLine("Shape has been sent to the back of the Z-order stack.");

            // Save the workbook
            string outputPath = "ZOrderDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
