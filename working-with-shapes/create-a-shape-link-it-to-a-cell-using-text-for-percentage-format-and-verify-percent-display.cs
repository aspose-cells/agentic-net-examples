using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Put a numeric value into cell A1 (25%)
                Cell cell = worksheet.Cells["A1"];
                cell.PutValue(0.25);

                // Apply percentage number format (built‑in format 10)
                Style style = cell.GetStyle();
                style.Number = 10; // percentage format
                cell.SetStyle(style);

                // Verify that the cell is recognized as a percent format
                Console.WriteLine("Cell A1 IsPercent: " + style.IsPercent); // should be True

                // Add a rectangle shape to the worksheet
                // Parameters: drawing type, upper left row, upper left column,
                // lower right row, lower right column, width, height
                Shape shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle,
                    2,   // upper left row
                    0,   // upper left column
                    5,   // lower right row
                    4,   // lower right column
                    100, // width
                    100  // height
                );

                // Optional display text for the shape
                shape.Text = "Linked Percent";

                // Link the shape to cell A1 (update and refresh enabled)
                shape.SetLinkedCell("A1", true, true);

                // Update the shape's displayed value from the linked cell
                shape.UpdateSelectedValue();

                // Save the workbook
                string outputPath = "ShapeLinkedPercent.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");

                // Read back the linked value from the shape (formatted text)
                string linkedValue = shape.Text;
                Console.WriteLine("Shape text after linking: " + linkedValue);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}