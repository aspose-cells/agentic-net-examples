// Title: Link a Rectangle Shape to a Bold Red Formatted Cell and Show Its Text – Aspose.Cells for .NET (C#)
// Description: Step‑by‑step C# example that creates a workbook, writes "Hello World" to B2, applies a bold red font style, adds a rectangle shape, links the shape to the styled cell, updates the shape’s displayed value, and prints verification details before saving the file.
// Keywords: Aspose.Cells | C# | .NET | shape linked cell | rectangle shape | UpdateSelectedValue | formatted cell text | bold red font | Excel shape text styling | link shape to cell Aspose
// Common Searches: Aspose.Cells link shape to cell with formatting | display cell style in linked shape .NET | rectangle shape shows bold red text Aspose.Cells | how to preserve cell formatting in linked shape | UpdateSelectedValue after styling cell
// Developer Intent: Create a rectangle shape, bind it to a styled cell, and verify that the shape displays the cell’s formatting.
// Use Cases: Dynamic dashboards where shapes reflect styled status labels from data cells. | Automated report templates that use linked shapes for highlighted headings. | Excel‑based UI elements that need to mirror bold, colored text defined in worksheet cells.
// AI Prompts: Show me C# code to link a rectangle shape to a cell and keep the cell’s bold red formatting visible in the shape using Aspose.Cells. | How can I verify that a shape displays the formatted text of its linked cell after applying a style in Aspose.Cells for .NET?

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Step‑by‑step C# example that creates a workbook, writes "Hello World" to B2, applies a bold red font style, adds a rectangle shape, links the shape to the styled cell, updates the shape’s displayed value, and prints verification details before saving the file.
class ShapeLinkedCellFormattedDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put formatted text into cell B2
            Cell linkedCell = worksheet.Cells["B2"];
            linkedCell.PutValue("Hello World");

            // Create a style with bold and red font and apply it to the cell
            Style style = workbook.CreateStyle();
            style.Font.IsBold = true;
            style.Font.Color = Color.Red;
            linkedCell.SetStyle(style);

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, row offset (pixels), column offset (pixels), width (pixels), height (pixels)
            Shape rectangle = worksheet.Shapes.AddRectangle(2, 1, 0, 0, 150, 40);

            // Link the shape's value to cell B2
            rectangle.SetLinkedCell("$B$2", false, false);

            // Update the shape's displayed text from the linked cell
            rectangle.UpdateSelectedValue();

            // Output verification information
            Console.WriteLine("Shape's LinkedCell: " + rectangle.LinkedCell);
            Console.WriteLine("Cell B2 value: " + linkedCell.StringValue);
            Console.WriteLine("Cell B2 font bold: " + linkedCell.GetStyle().Font.IsBold);
            Console.WriteLine("Cell B2 font color (ARGB): " + linkedCell.GetStyle().Font.Color.ToArgb());

            // Ensure the output directory exists
            string outputPath = "ShapeLinkedCellFormatted.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + Path.GetFullPath(outputPath));
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
