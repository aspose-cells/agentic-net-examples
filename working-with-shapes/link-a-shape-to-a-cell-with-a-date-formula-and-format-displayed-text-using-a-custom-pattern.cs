// Title: How to link a rectangle shape to a TODAY() formula cell and display a custom formatted date using Aspose.Cells for .NET (C#)
// AI Prompts: Create C# code that adds a rectangular drawing object, associates its displayed text with the value of a cell that uses the current date function, and formats the text as dd-MMM-yyyy using Aspose.Cells. | Generate an Excel workbook with Aspose.Cells where a shape automatically reflects the formatted date from a linked cell and save the file as LinkedShape.xlsx.
// Common Searches: Aspose.Cells C# shape text bound to TODAY() cell value | display custom formatted date inside a rectangle shape using Aspose.Cells | how to update shape text automatically after formula recalculation in Aspose.Cells
// Tags: link shape to cell Aspose.Cells C# | add shape type Rectangle Aspose.Cells | custom date format dd-MMM-yyyy Aspose.Cells | bind shape text to formula result Aspose.Cells | save workbook with linked shape Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, inserts a TODAY() formula in cell B2, applies a custom date format (dd-MMM-yyyy), recalculates formulas, adds a rectangle shape, sets the shape's text to the formatted cell value, and saves the workbook as LinkedShape.xlsx.
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

            // Put a date formula in cell B2
            Cell dateCell = sheet.Cells["B2"];
            dateCell.Formula = "=TODAY()";

            // Apply a custom date format pattern (e.g., 15-Mar-2023)
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "dd-MMM-yyyy";
            dateCell.SetStyle(dateStyle);

            // Recalculate formulas so the date value is available
            workbook.CalculateFormula();

            // Add a rectangle shape to the worksheet
            // Parameters: type, upperLeftRow, upperLeftColumn, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                5,    // upper left row
                0,    // upper left column
                0,    // top offset (in pixels)
                0,    // left offset (in pixels)
                30,   // height (in points)
                150); // width (in points)

            // Set initial text (will be refreshed automatically if linked)
            shape.Text = dateCell.StringValue;

            // Define output file path
            string outputPath = "LinkedShape.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
