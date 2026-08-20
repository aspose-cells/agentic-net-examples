// Title: C# Example: Add Rectangle Shape Linked to a Cell with Custom Thousand‑Separator Number Format using Aspose.Cells
// Description: This sample creates a new workbook, writes the value 1234567 to cell B2, applies the custom number format "#,##0" (thousand separator), adds a rectangle shape, links the shape to B2 using A1 notation, and saves the file as ShapeLinkedNumberFormat.xlsx. The shape automatically shows the formatted cell value.
// Keywords: Aspose.Cells | C# | .NET | add rectangle shape | linked cell | custom number format | thousand separator | Excel shape linking | sample code | example
// Common Searches: Aspose.Cells link shape to cell C# | how to apply thousand separator format in Aspose.Cells | add rectangle shape and bind to cell using Aspose.Cells .NET | shape linked cell custom number format Aspose.Cells | C# example for linking shape to worksheet cell
// Developer Intent: Create a shape that displays a cell's value with a custom thousand‑separator format and stays synchronized with the cell.
// Use Cases: Financial dashboards where a shape shows a total amount formatted with commas. | Invoice templates that display the billed amount inside a shape, automatically updating when the cell changes. | Interactive reports that use linked shapes to highlight key metrics with custom numeric formatting.
// AI Prompts: Generate C# code that adds a rectangle shape, links it to a specific cell, and applies the "#,##0" number format using Aspose.Cells for .NET. | Explain how to set a shape's linked cell so the shape displays the cell's formatted value and updates automatically when the cell value changes.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkExample
{
    // This sample creates a new workbook, writes the value 1234567 to cell B2, applies the custom number format "#,##0" (thousand separator), adds a rectangle shape, links the shape to B2 using A1 notation, and saves the file as ShapeLinkedNumberFormat.xlsx. The shape automatically shows the formatted cell value.
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
                Cell targetCell = sheet.Cells["B2"];
                targetCell.PutValue(1234567);

                // Create a custom style with thousand separator format "#,##0"
                Style customStyle = workbook.CreateStyle();
                customStyle.Custom = "#,##0";

                // Apply the custom style to the target cell
                targetCell.SetStyle(customStyle);

                // Add a rectangle shape to the worksheet
                // Parameters: shape type, upper left row, upper left column, top, left, width, height
                Shape rectShape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 4, 0, 0, 0, 150, 50);

                // Link the shape's value to cell B2 (A1 style, locale‑aware)
                rectShape.SetLinkedCell("$B$2", false, true);

                // Optionally set some text for the shape (will display the linked value)
                rectShape.Text = "Linked Value";

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
}
