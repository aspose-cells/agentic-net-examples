// Title: Add a rectangle shape with custom text to an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a rectangle shape on a worksheet canvas, assigns the text "Custom Rectangle Text", applies blue 12‑point font, and saves the file as an .xlsx with Aspose.Cells. | Write a method that inserts a rectangle at row 3, column B with specific pixel offsets, sets its caption, and returns the modified Workbook object. | Provide a script to modify an existing workbook by adding a positioned rectangle shape, customizing its text and font, and persisting the changes.
// Common Searches: Aspose.Cells C# how to insert a rectangle shape with text into a specific cell | set custom caption for a rectangle drawing in an Excel file using Aspose.Cells | position a shape on the worksheet canvas by row and column with Aspose.Cells .NET | change font color and size of shape text in Aspose.Cells workbook | save workbook after adding shapes with Aspose.Cells for .NET
// Tags: add rectangle shape Aspose.Cells C# | set shape text Aspose.Cells | position shape on worksheet canvas | format shape font Aspose.Cells | save workbook with shapes Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, adds a rectangle shape at a defined cell location, sets custom text with blue 12‑point font, and saves the result as RectangleShape.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define shape position and size (in pixels)
        int upperLeftRow = 2;          // Row index (0‑based)
        int upperLeftColumn = 1;       // Column index (0‑based)
        int top = 5;                   // Top offset within the cell
        int left = 5;                  // Left offset within the cell
        int height = 100;              // Height of the rectangle
        int width = 200;               // Width of the rectangle

        // Add a rectangle shape to the worksheet canvas
        Shape rectangle = sheet.Shapes.AddShape(
            MsoDrawingType.Rectangle,
            upperLeftRow,
            upperLeftColumn,
            top,
            left,
            height,
            width);

        // Set custom text for the rectangle
        rectangle.Text = "Custom Rectangle Text";

        // Optional: format the text
        rectangle.Font.Color = Color.Blue;
        rectangle.Font.Size = 12;

        // Save the workbook to a file
        workbook.Save("RectangleShape.xlsx");
    }
}
