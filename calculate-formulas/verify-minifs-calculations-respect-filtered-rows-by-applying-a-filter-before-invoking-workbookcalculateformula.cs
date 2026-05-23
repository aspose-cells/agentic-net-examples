using System;
using Aspose.Cells;

namespace AsposeCellsMinifsFilterDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Populate sample data
            // Header row
            cells["A1"].PutValue("Value");
            cells["B1"].PutValue("Category");

            // Rows 2-11: Values 1-10, Category alternates between "X" and "Y"
            for (int i = 0; i < 10; i++)
            {
                int row = i + 1; // zero‑based index (row 2 in Excel)
                cells[row, 0].PutValue(i + 1);                     // Column A: 1..10
                cells[row, 1].PutValue((i % 2 == 0) ? "X" : "Y"); // Column B: X,Y alternating
            }

            // 3. Insert MINIFS formula in C1:
            // =MINIFS(A2:A11, B2:B11, "X")
            cells["C1"].Formula = "=MINIFS(A2:A11,B2:B11,\"X\")";

            // 4. Calculate formulas without any filter
            workbook.CalculateFormula();

            // Capture the result before filtering
            double resultWithoutFilter = cells["C1"].DoubleValue;
            Console.WriteLine($"MINIFS result without filter: {resultWithoutFilter}");

            // 5. Apply an AutoFilter to hide rows where Category = "X"
            // Set the filter range to include headers and data rows
            sheet.AutoFilter.SetRange(0, 0, 10); // rows 0‑10 (A1:B11)

            // Filter column B (field index 1) for "Y" – this will hide all "X" rows
            sheet.AutoFilter.Filter(1, "Y");
            // Refresh to apply the filter
            sheet.AutoFilter.Refresh();

            // 6. Re‑calculate formulas after the filter is applied
            workbook.CalculateFormula();

            // Capture the result after filtering
            double resultWithFilter = cells["C1"].DoubleValue;
            Console.WriteLine($"MINIFS result after filtering out \"X\" rows: {resultWithFilter}");

            // 7. Verify that the filtered result respects hidden rows
            // Expected: since no visible rows meet the criteria "X", MINIFS should return 0
            if (Math.Abs(resultWithFilter) < 1e-9)
            {
                Console.WriteLine("Verification passed: MINIFS ignored hidden rows.");
            }
            else
            {
                Console.WriteLine("Verification failed: MINIFS did not respect the filter.");
            }

            // 8. Save the workbook for visual inspection (optional)
            workbook.Save("MinifsFilterDemo.xlsx");
        }
    }
}