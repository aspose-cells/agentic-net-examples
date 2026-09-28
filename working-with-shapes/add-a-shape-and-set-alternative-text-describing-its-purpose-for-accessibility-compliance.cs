// Title: Insert a rectangle shape at cell B2 and set its alternative text for accessibility with Aspose.Cells in C#
// AI Prompts: Generate C# code that creates a workbook, adds a rectangle shape anchored at B2, assigns an AlternativeText describing its purpose, and saves the file as XLSX using Aspose.Cells. | Write a C# snippet that uses Aspose.Cells to place a rectangle shape on the first worksheet, set its AlternativeText property for screen‑reader accessibility, and export the workbook. | Provide a C# example demonstrating how to add a shape to an Excel sheet with Aspose.Cells and configure the shape's AlternativeText before saving.
// Common Searches: how to set alternative text for a shape in Aspose.Cells C# | Aspose.Cells add rectangle shape at specific cell with alt text | C# Aspose.Cells shape accessibility example | set shape AlternativeText property in Excel using Aspose.Cells .NET | add shape to worksheet and define alt text for screen readers Aspose.Cells
// Tags: Aspose.Cells add rectangle shape C# | Aspose.Cells shape alternative text .NET | Excel shape accessibility Aspose.Cells | C# insert shape with alt text Aspose.Cells | Aspose.Cells shape property AlternativeText

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, adds a rectangle shape anchored at B2, assigns alternative text for accessibility, and saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape (row 2, column 2) with width 150 and height 80 points
            // AddShape returns a Shape object in recent Aspose.Cells versions
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 150, 80);

            // Set alternative text for accessibility compliance
            shape.AlternativeText = "Rectangle indicating the sales target area.";

            // Save the workbook
            workbook.Save("Result.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
