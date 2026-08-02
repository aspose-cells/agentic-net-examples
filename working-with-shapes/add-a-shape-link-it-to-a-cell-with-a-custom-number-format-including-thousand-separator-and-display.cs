// Title: Link a Rectangle Shape to a Formatted Cell with Thousand Separator Using Aspose.Cells for .NET
// Description: Shows how to put a number in B2, apply the "#,##0.00" format, add a rectangle shape, bind it to the cell, and save the workbook with Aspose.Cells C#.
// Keywords: Aspose.Cells | C# Excel shape linking | custom numeric format | thousand separator | rectangle shape | SetLinkedCell method | Excel automation .NET | formatted cell display | Aspose.Cells example | shape value binding
// Common Searches: Aspose.Cells link shape to cell example | C# add rectangle shape to worksheet | apply thousand separator format with Aspose.Cells | display formatted cell value in a shape | SetLinkedCell usage in .NET
// Developer Intent: Create a shape that reflects a cell’s formatted numeric value.
// Use Cases: Financial dashboards where a shape shows totals with commas for readability. | Invoice templates that display monetary amounts inside shapes using a custom format. | Interactive reports that update shape captions automatically when underlying cell values change.
// AI Prompts: Generate C# code that adds an oval shape linked to cell C5 and formats the value as currency "#,##0.00" with Aspose.Cells. | Explain how to refresh linked shapes after modifying cell data and re‑saving the workbook in Aspose.Cells. | Provide a sample that links three different shapes to three cells, each using a distinct custom number format, using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to put a number in B2, apply the "#,##0.00" format, add a rectangle shape, bind it to the cell, and save the workbook with Aspose.Cells C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a numeric value into cell B2
            Cell cell = sheet.Cells["B2"];
            cell.PutValue(1234567.89);

            // Create a custom style with thousand separator and two decimal places
            Style style = workbook.CreateStyle();
            style.Custom = "#,##0.00";

            // Apply the custom style to the cell
            cell.SetStyle(style);

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top, left, width, height
            Shape rectangle = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 150, 20, 200, 50);

            // Link the shape's value to cell B2
            rectangle.SetLinkedCell("$B$2", false, true);

            // Define output file path
            string outputPath = "ShapeLinkedNumberFormat.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
