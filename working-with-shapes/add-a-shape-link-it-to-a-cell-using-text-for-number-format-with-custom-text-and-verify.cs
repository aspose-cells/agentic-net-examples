using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeLinkedCellExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set a numeric value in cell B2
            Cell linkedCell = sheet.Cells["B2"];
            linkedCell.PutValue(12345);

            // Apply a custom number format that includes custom text
            // Example format: "Total: " followed by the number
            Style style = linkedCell.GetStyle();
            style.Custom = "\"Total: \"0";
            linkedCell.SetStyle(style);

            // Add a label shape (text box) to the worksheet
            // Note: In some Aspose.Cells versions AddLabel requires six integer parameters.
            // The last two parameters define the lower‑right cell of the shape.
            Label label = sheet.Shapes.AddLabel(1, 1, 50, 200, 1, 1);
            label.Text = "Linked Value:";   // Set the displayed text

            // Link the shape to the cell B2
            label.LinkedCell = "$B$2";

            // Verify the linked cell address
            Console.WriteLine("Shape's LinkedCell: " + label.LinkedCell);

            // Refresh the shape's displayed value based on the linked cell
            label.UpdateSelectedValue();

            // Define output file path
            string outputPath = "ShapeLinkedCellExample.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}