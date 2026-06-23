using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var sheet = workbook.Worksheets[0];

                // Put a sample value in cell A1 (used by the IF formula)
                sheet.Cells["A1"].PutValue(15);

                // Set an IF formula in cell A2 that will be linked to the shape
                sheet.Cells["A2"].Formula = "=IF(A1>10,\"High\",\"Low\")";

                // Add a label shape (row 4, column 2, height 30, width 100, top offset 0, left offset 0)
                var label = sheet.Shapes.AddLabel(4, 2, 30, 100, 0, 0);

                // Link the label's displayed text to the cell containing the IF formula (A2)
                label.SetLinkedCell("A2", false, false);

                // Refresh the shape so it shows the current value of the linked cell
                label.UpdateSelectedValue();

                // Define output file path
                string outputPath = "ShapeWithConditionalText.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}