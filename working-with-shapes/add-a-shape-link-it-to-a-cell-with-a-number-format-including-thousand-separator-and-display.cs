// Title: Create a rectangle shape linked to a formatted cell with thousand separators using Aspose.Cells for .NET
// AI Prompts: Add a rectangle shape to a worksheet and set its Text property to a formula that references cell A1 so the shape shows the cell’s formatted number. | Apply the custom number format "#,##0" to a cell, then save the workbook, ensuring the linked shape updates automatically when the cell value changes.
// Common Searches: Aspose.Cells C# bind shape text to a cell value with custom number formatting | link rectangle shape to cell A1 and display thousand separator format in Excel using Aspose.Cells | show formatted numeric cell value inside a shape with Aspose.Cells for .NET | automatically update shape text when the linked cell changes Aspose.Cells example
// Tags: add rectangle shape Aspose.Cells | shape text formula binding Aspose.Cells | apply thousand separator format C# | link shape to cell value Aspose.Cells | save workbook with linked shape .xlsx

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, writes 1234567 to cell A1, applies the "#,##0" thousand‑separator format, adds a rectangle shape, links the shape’s text to A1 with a formula so the formatted value appears, centers the text, sets the font size, and saves the file as ShapeLinkedToCell.xlsx.
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

            // Set a numeric value in cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(1234567);

            // Apply number format with thousand separator (e.g., 1,234,567)
            Style style = cell.GetStyle();
            style.Custom = "#,##0";
            cell.SetStyle(style);

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 60, 200);

            // Link the shape's text to cell A1 using a formula.
            // The shape will display the formatted value of A1 and update automatically when A1 changes.
            shape.Text = "=A1";

            // Optional: format the shape's text (font size, alignment, etc.)
            shape.TextHorizontalAlignment = TextAlignmentType.Center;
            shape.TextVerticalAlignment = TextAlignmentType.Center;
            shape.Font.Size = 12;

            // Determine output path and ensure the directory exists
            string outputPath = "ShapeLinkedToCell.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
