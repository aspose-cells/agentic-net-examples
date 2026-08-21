// Title: Add Linked Rectangle Shapes to All Worksheets with Aspose.Cells for .NET
// Description: Creates a new Workbook, adds extra worksheets, writes a summary to cell A1 on each sheet, then iterates through every worksheet to insert a rectangle shape at row 2/column 2, assigns a unique name, links the shape to the sheet’s A1 cell, applies light‑blue fill and dark‑blue border, logs the shape details, and saves the file as WorksheetsWithLinkedShapes.xlsx.
// Keywords: Aspose.Cells add shape | link shape to cell | iterate worksheets C# | rectangle shape Aspose.Cells | SetLinkedCell example | shape logging Aspose | batch shape insertion | Excel navigation button
// Common Searches: How to add a rectangle shape to every worksheet with Aspose.Cells | Aspose.Cells set linked cell for a shape in C# | Iterate all worksheets and add named shapes | Aspose.Cells shape formatting and linking example | C# code to create navigation shapes in Excel workbook
// Developer Intent: Insert a rectangle shape on each worksheet, link it to the sheet’s A1 cell, and output the shape’s name, type, and linked cell.
// Use Cases: Build a consistent navigation button on each sheet that jumps to a summary section. | Automate branding elements (shapes with corporate colors) across a multi‑sheet report. | Generate a visual dashboard where each worksheet contains a clickable shape linked to its key data cell.
// AI Prompts: Generate C# code using Aspose.Cells that adds a circle shape to every worksheet, links it to cell B2, sets a green fill, and prints the linked cell address. | Show an example that loops through all worksheets, inserts a rectangle shape, gives each shape a unique name, links it to a specific cell, customizes its colors, and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkDemo
{
    // Creates a new Workbook, adds extra worksheets, writes a summary to cell A1 on each sheet, then iterates through every worksheet to insert a rectangle shape at row 2/column 2, assigns a unique name, links the shape to the sheet’s A1 cell, applies light‑blue fill and dark‑blue border, logs the shape details, and saves the file as WorksheetsWithLinkedShapes.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (default worksheet is already added)
            Workbook workbook = new Workbook();

            // Add a couple of extra worksheets for demonstration
            Worksheet ws1 = workbook.Worksheets.Add("DataSheet1");
            Worksheet ws2 = workbook.Worksheets.Add("DataSheet2");

            // Populate a summary cell in each worksheet (e.g., A1)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.Cells["A1"].PutValue($"Summary of {sheet.Name}");
            }

            // Iterate through all worksheets, add a rectangle shape, link it to the summary cell, and log details
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Add a rectangle shape (row, top, column, left, height, width)
                // Placing it at row 2, column 2 with size 100x50 points
                Shape shape = sheet.Shapes.AddRectangle(2, 0, 2, 0, 100, 50);

                // Give the shape a distinctive name
                shape.Name = $"SummaryShape_{sheet.Index}";

                // Link the shape to the summary cell (A1) of the current worksheet
                // Parameters: cell name, isRowAbsolute, isColumnAbsolute
                shape.SetLinkedCell("A1", false, false);

                // Optionally set some visual properties
                shape.FillFormat.ForeColor = System.Drawing.Color.LightBlue;
                shape.LineFormat.ForeColor = System.Drawing.Color.DarkBlue;

                // Log the result
                string linkedCell = shape.GetLinkedCell(false, false);
                Console.WriteLine($"Worksheet '{sheet.Name}' (Index {sheet.Index}): Added shape '{shape.Name}' of type {shape.Type} linked to cell '{linkedCell}'.");
            }

            // Save the workbook to a file
            workbook.Save("WorksheetsWithLinkedShapes.xlsx");
        }
    }
}
