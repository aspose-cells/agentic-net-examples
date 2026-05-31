using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

class RunningTotalInTable
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ----- Populate sample data -----
            // Header row
            cells["A1"].PutValue("Amount");
            cells["B1"].PutValue("Running Total");

            // Sample amounts (rows 2‑6)
            double[] amounts = { 100, 150, 200, 250, 300 };
            for (int i = 0; i < amounts.Length; i++)
            {
                cells[i + 1, 0].PutValue(amounts[i]); // Column A
            }

            // ----- Create a table that includes both columns -----
            // Table range: A1:B6 (header + 5 data rows)
            int tableIndex = sheet.ListObjects.Add(0, 0, amounts.Length, 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.ShowTotals = false; // totals row not required

            // ----- Set running total formula for each data row -----
            // Use A1‑style formulas referencing the worksheet cells directly.
            for (int i = 0; i < amounts.Length; i++)
            {
                int excelRow = i + 2; // Excel rows start at 1; row 2 is first data row
                string formula = $"=A{excelRow}+IFERROR(B{excelRow - 1},0)";
                cells[excelRow - 1, 1].Formula = formula; // Column B (index 1)
            }

            // Calculate all formulas so the running totals are materialized
            workbook.CalculateFormula();

            // ----- Save the workbook -----
            string outputPath = "RunningTotalTable.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}