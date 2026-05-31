using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeLinkedArrayConstantDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define an array constant and set it as an array formula in cell A1
            // The formula "{1,2,3}" will spill into three adjacent cells (A1:C1)
            Cell arrayCell = cells["A1"];
            arrayCell.SetArrayFormula("{1,2,3}", 1, 3); // 1 row, 3 columns

            // Add a ListBox shape to the worksheet (positional overload)
            // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, width, height
            ListBox listBox = worksheet.Shapes.AddListBox(
                2,   // upper left row (zero‑based)
                0,   // upper left column (zero‑based)
                2,   // row offset in pixels
                0,   // column offset in pixels
                130, // width in pixels
                30   // height in pixels
            );

            // Link the ListBox value to the cell containing the array constant (A1)
            listBox.SetLinkedCell("$A$1", isR1C1: false, isLocal: false);

            // Update the ListBox selected value based on the linked cell
            worksheet.Shapes.UpdateSelectedValue();

            // Output the linked cell address and the currently selected index
            Console.WriteLine("LinkedCell: " + listBox.LinkedCell);
            Console.WriteLine("SelectedIndex: " + listBox.SelectedIndex);

            // Save the workbook to a file (ensure the directory exists)
            string outputPath = "ShapeLinkedArrayConstant.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}