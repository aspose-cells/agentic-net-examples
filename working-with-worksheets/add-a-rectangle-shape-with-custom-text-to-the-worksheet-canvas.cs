// Title: Create and format a rectangle shape with custom text on an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds a rectangle shape to a worksheet canvas, sets its inner text, applies a light blue fill and dark blue solid border, and saves the workbook as an .xlsx file with Aspose.Cells. | Write a C# snippet to position a rectangle shape at a specific row and column, define its size, customize fill and line formatting, embed custom text, and export the workbook using Aspose.Cells.
// Common Searches: Aspose.Cells C# add rectangle shape with text to specific cell range | how to set fill color and border style for a shape in Aspose.Cells .NET | positioning a rectangle shape on an Excel worksheet using Aspose.Cells API | save workbook after inserting custom shaped annotation with Aspose.Cells
// Tags: add rectangle shape Aspose.Cells C# | set shape text Aspose.Cells | shape fill color Aspose.Cells .NET | shape border formatting Aspose.Cells | worksheet canvas shape positioning Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Drawing;

// Demonstrates how to create a new workbook, add a rectangle shape to the first worksheet at a specified cell location, assign custom text, apply a light‑blue fill and dark‑blue solid border with custom line weight, and save the file as RectangleShape.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define shape position and size (in pixels)
            int upperLeftRow = 2;      // Row index where the shape starts
            int upperLeftColumn = 2;   // Column index where the shape starts
            int top = 5;               // Distance from the top of the cell (pixels)
            int left = 5;              // Distance from the left of the cell (pixels)
            int height = 100;          // Height of the shape (pixels)
            int width = 200;           // Width of the shape (pixels)

            // Add a rectangle shape to the worksheet canvas
            Shape rectangle = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                upperLeftRow,
                upperLeftColumn,
                top,
                left,
                height,
                width);

            // Set custom text inside the rectangle
            rectangle.Text = "Custom Rectangle Text";

            // Optional formatting
            rectangle.FillFormat.ForeColor = Color.LightBlue;               // Background color
            rectangle.LineFormat.Weight = 1.5;                              // Border thickness
            rectangle.LineFormat.DashStyle = MsoLineDashStyle.Solid;        // Border style
            rectangle.LineFormat.ForeColor = Color.DarkBlue;                // Border color

            // Save the workbook to a file
            string outputPath = "RectangleShape.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
