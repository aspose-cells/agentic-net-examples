// Title: Create a rectangle shape linked to a custom‑formatted cell and verify the displayed currency text with Aspose.Cells for .NET
// AI Prompts: Insert a rectangle drawing on the first worksheet, bind it to cell A1 that uses a custom number format, and print the shape's Text property. | Add an assertion that the shape's Text matches the formatted value of the linked cell (e.g., "$1,234.57"). | Save the workbook as an .xlsx file after linking the shape and display the shape's text in the console.
// Common Searches: Aspose.Cells how to bind a shape to a cell with a custom number format | C# example linking rectangle shape to a formatted cell and reading shape text | verify that a linked shape displays the currency format of its source cell in Aspose.Cells | retrieve formatted cell value from a shape's Text property using Aspose.Cells for .NET
// Tags: rectangle shape linked cell Aspose.Cells | cell formatted as currency Aspose.Cells | shape text shows formatted value | link shape to formatted cell .NET | retrieve displayed text from linked shape

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates creating a workbook, applying a custom currency number format to cell A1, adding a rectangle shape, linking the shape to that cell, reading the shape's Text (which reflects the formatted value), and saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set a custom number format in cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(1234.567);
            Style style = cell.GetStyle();
            style.Custom = "$#,##0.00"; // Currency format with two decimals
            cell.SetStyle(style);

            // Add a rectangle shape to the worksheet
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // Shape type
                5,   // Upper left row
                0,   // Upper left row offset (pixels)
                5,   // Upper left column
                0,   // Upper left column offset (pixels)
                30,  // Height (pixels)
                100  // Width (pixels)
            );

            // Link the shape to cell A1 so it displays the formatted value
            shape.LinkedCell = "A1";

            // Retrieve and display the shape's text (should reflect the cell's format)
            string displayedText = shape.Text; // Expected: "$1,234.57"
            Console.WriteLine("Shape text: " + displayedText);

            // Save the workbook
            string outputPath = "ShapeLinkedCell.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
