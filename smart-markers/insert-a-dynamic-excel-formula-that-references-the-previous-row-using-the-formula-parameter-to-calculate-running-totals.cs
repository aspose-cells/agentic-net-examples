// Title: Add a cumulative running‑total column to an Aspose.Cells ListObject with a formula that references the previous row (C#)
// AI Prompts: Write C# code using Aspose.Cells to create a ListObject and set a formula in the third column that adds the current Amount cell to the Running Total cell of the previous row. | Show how to programmatically apply a dynamic Excel formula for cumulative sums to each data row of a table in Aspose.Cells. | Generate a complete C# example that builds a workbook, fills sample data, inserts a running‑total formula referencing the prior row, calculates formulas, and saves the file.
// Common Searches: aspnet c# how to calculate running total in an Aspose.Cells table using previous row reference | aspose.cells cumulative sum formula previous row example | insert dynamic formula into ListObject column c# aspose.cells | calculate running total column in Excel workbook with Aspose.Cells API
// Tags: Aspose.Cells ListObject running total formula | C# cumulative sum formula in Excel table | dynamic previous‑row reference Aspose.Cells | populate table column with formula Aspose.Cells | calculate running totals programmatically C#

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsRunningTotalDemo
{
    // The sample creates a workbook, adds headers and sample data, defines a ListObject covering columns A‑C, and inserts a running‑total formula into the third column that references the previous row (C of previous row + B of current row). After calculating formulas, the workbook is saved as RunningTotalDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["C1"].PutValue("Running Total");

            // Sample data (ID, Amount)
            int[] ids = { 1, 2, 3, 4, 5 };
            double[] amounts = { 10, 20, 15, 30, 25 };

            // Populate the data rows (starting at row 2)
            for (int i = 0; i < ids.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(ids[i]);      // Column A
                sheet.Cells[i + 1, 1].PutValue(amounts[i]); // Column B
            }

            // Create a ListObject (table) that covers the data range including the header
            // Table range: A1:C{lastRow}
            int lastRow = ids.Length + 1; // +1 for header
            int tableIndex = sheet.ListObjects.Add(0, 0, lastRow, 2, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Insert running‑total formula into the "Running Total" column (column offset 2)
            // Row offset 0 = header, so data rows start at offset 1
            for (int rowOffset = 1; rowOffset <= ids.Length; rowOffset++)
            {
                // Corresponding worksheet row number (1‑based)
                int worksheetRow = rowOffset + 1; // because header is row 1

                string formula;
                if (rowOffset == 1)
                {
                    // First data row: running total equals the amount itself
                    formula = $"=B{worksheetRow}";
                }
                else
                {
                    // Subsequent rows: previous total (C of previous row) + current amount (B of this row)
                    formula = $"=C{worksheetRow - 1}+B{worksheetRow}";
                }

                // Apply the formula to the cell in the table
                table.PutCellFormula(rowOffset, 2, formula);
            }

            // Calculate all formulas so that the running totals are materialized
            workbook.CalculateFormula();

            // Save the workbook
            workbook.Save("RunningTotalDemo.xlsx");
        }
    }
}
