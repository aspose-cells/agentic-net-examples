using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a date formula in cell B2
            Cell dateCell = sheet.Cells["B2"];
            dateCell.Formula = "=TODAY()";

            // Apply a custom date format (e.g., 25-Dec-2023)
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "dd-mmm-yyyy";

            // Apply the style to the cell (all style attributes)
            dateCell.SetStyle(dateStyle);

            // Add a rectangle shape to the sheet
            // Parameters: upper row, left column, upper offset, left offset, width, height (in pixels)
            RectangleShape shape = sheet.Shapes.AddRectangle(4, 1, 10, 10, 150, 50);

            // Link the shape to the date cell (A1‑style reference, locale‑aware)
            shape.SetLinkedCell("$B$2", false, true);

            // Format the displayed text of the shape (font size and color)
            // Using (0,0) formats all characters in the shape
            shape.Characters(0, 0).Font.Size = 14;
            shape.Characters(0, 0).Font.Color = Color.Blue;

            // Determine output path and ensure its directory exists
            string outputPath = "LinkedShapeDate.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}