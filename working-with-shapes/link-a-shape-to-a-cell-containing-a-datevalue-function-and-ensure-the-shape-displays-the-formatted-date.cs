// Title: Link a Label Shape to a DATEVALUE Cell and Display Formatted Date – Aspose.Cells for .NET (C#)
// Description: Creates a workbook, inserts a DATEVALUE formula in A1, applies a short‑date number format, adds a label shape, links it to the cell via the LinkedCell property, forces an immediate refresh with UpdateSelectedValue, and saves the file.
// Keywords: Aspose.Cells | C# | label shape | LinkedCell | DATEVALUE | date formatting | shape linking | Excel automation | short date format
// Common Searches: Aspose.Cells link shape to cell | C# label shape display date | LinkedCell property example | format date in linked shape Aspose.Cells | refresh shape after changing cell formula Aspose.Cells
// Developer Intent: Bind a shape to a cell containing a DATEVALUE formula and have the shape show the cell’s formatted date.
// Use Cases: Dynamic dashboard date stamps that update automatically when the source formula changes. | Report headers that reflect a calculated start date while preserving short‑date formatting. | Printable invoices where a shape displays the invoice date derived from a DATEVALUE function.
// AI Prompts: Generate C# code using Aspose.Cells to add a label shape, link it to a DATEVALUE cell, and show the formatted date. | Explain how to refresh a linked shape after modifying the cell’s formula or number format in Aspose.Cells for .NET. | Provide a sample that links multiple label shapes to different date cells, each preserving its own date format.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsLinkShapeToDate
{
    // Creates a workbook, inserts a DATEVALUE formula in A1, applies a short‑date number format, adds a label shape, links it to the cell via the LinkedCell property, forces an immediate refresh with UpdateSelectedValue, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a DATEVALUE formula in cell A1
            Cell dateCell = sheet.Cells["A1"];
            dateCell.Formula = "=DATEVALUE(\"2023-01-01\")";

            // Apply a date number format to display the serial number as a date
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // Built‑in short date format
            dateCell.SetStyle(dateStyle);

            // Add a label shape that will show the linked cell's value
            // Parameters: upper left row, upper left column, top offset, left offset, height, width
            Label label = (Label)sheet.Shapes.AddLabel(2, 2, 0, 0, 30, 150);

            // Link the label to cell A1 so it displays the formatted date
            // Using the property setter (absolute A1 reference)
            label.LinkedCell = "$A$1";

            // Optionally, update the shape's displayed value immediately
            label.UpdateSelectedValue();

            // Save the workbook
            workbook.Save("ShapeLinkedToDate.xlsx");
        }
    }
}
