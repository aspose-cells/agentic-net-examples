using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeLinkedDateExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set a date value in cell B2 and apply a custom date format (e.g., "dd-MMM-yyyy")
            Cell dateCell = sheet.Cells["B2"];
            dateCell.PutValue(DateTime.Now);
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "dd-MMM-yyyy";
            dateCell.SetStyle(dateStyle);

            // Add a rectangle shape to the worksheet (position: row 4, column 2, width 100, height 30)
            // MsoDrawingType.Rectangle is used for a simple rectangle shape
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 4, 2, 0, 0, 100, 30);

            // Set the shape's text to display the linked date using the same custom format
            shape.Text = "Date: " + dateCell.StringValue;

            // Link the shape to the cell B2 so that its value can be updated automatically
            // The formula is in A1‑style notation, not R1C1, and locale‑aware (true)
            shape.SetLinkedCell("$B$2", false, true);

            // Verify that the linked cell contains a DateTime value and matches the custom format
            object linkedValue = sheet.Cells["B2"].Value;
            if (linkedValue is DateTime dt)
            {
                string formatted = dt.ToString("dd-MMM-yyyy");
                Console.WriteLine("Linked cell contains a date: " + formatted);
            }
            else
            {
                Console.WriteLine("Linked cell does not contain a date.");
            }

            // Save the workbook to a file
            string outputPath = "ShapeLinkedDate.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}