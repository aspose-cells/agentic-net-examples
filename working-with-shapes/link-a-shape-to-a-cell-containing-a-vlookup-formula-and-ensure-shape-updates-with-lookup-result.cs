using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate a lookup table (A2:B5)
            worksheet.Cells["A2"].Value = "Apple";
            worksheet.Cells["B2"].Value = 1.2;
            worksheet.Cells["A3"].Value = "Banana";
            worksheet.Cells["B3"].Value = 0.8;
            worksheet.Cells["A4"].Value = "Cherry";
            worksheet.Cells["B4"].Value = 2.5;
            worksheet.Cells["A5"].Value = "Date";
            worksheet.Cells["B5"].Value = 3.0;

            // Cell C2 holds the lookup key, D2 will hold the VLOOKUP result
            worksheet.Cells["C2"].Value = "Apple";
            worksheet.Cells["D2"].Formula = "=VLOOKUP(C2,$A$2:$B$5,2,FALSE)";

            // Add a label shape that will display the VLOOKUP result
            // Parameters: upperLeftRow, upperLeftColumn, height, width, shapeIndex, isVertical (0 = false)
            Shape labelShape = worksheet.Shapes.AddLabel(2, 2, 30, 100, 0, 0);

            // Link the label to the cell containing the VLOOKUP result
            // isR1C1 = false (A1 style), isLocal = true (locale aware)
            labelShape.SetLinkedCell("$D$2", false, true);

            // Calculate formulas so D2 gets a value
            workbook.CalculateFormula();

            // Update the shape to reflect the linked cell's current value
            labelShape.UpdateSelectedValue();

            // Change the lookup key to demonstrate dynamic update
            worksheet.Cells["C2"].Value = "Cherry";
            workbook.CalculateFormula();
            labelShape.UpdateSelectedValue();

            // Save the workbook
            string outputPath = "LinkedShapeVlookup.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}