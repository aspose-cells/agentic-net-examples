using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class MergeAndAverageDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a range (A1:B2), assign a name, and merge its cells
            AsposeRange range = worksheet.Cells.CreateRange("A1", "B2");
            range.Name = "MyRange";               // optional named range
            range.Merge();                        // combine cells into a single cell

            // Populate data for the average calculation (C1:C5)
            for (int i = 0; i < 5; i++)
            {
                worksheet.Cells[i, 2].PutValue(i + 1); // C1..C5 = 1,2,3,4,5
            }

            // Set the formula in the merged cell (upper‑left cell of the range)
            Cell mergedCell = worksheet.Cells["A1"];
            mergedCell.Formula = "=AVERAGE(C1:C5)";

            // Recalculate formulas so the result is updated
            workbook.CalculateFormula();

            // Save the workbook
            string outputPath = "MergedAverage.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}