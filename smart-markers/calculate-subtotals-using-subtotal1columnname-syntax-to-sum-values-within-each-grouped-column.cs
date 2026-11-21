// Title: Insert grouped subtotal rows in an Excel file using Aspose.Cells SUBTOTAL function with (subtotal1:ColumnName) syntax in C#
// Common Searches: how to add subtotal rows for each category using Aspose.Cells in C# | Aspose.Cells C# example for generating SUBTOTAL(9) formulas with dynamic ranges | C# code to insert subtotal rows based on a grouping column in an Excel worksheet | using (subtotal1:ColumnName) syntax in Aspose.Cells to calculate group totals | Aspose.Cells tutorial for auto‑fitting columns after inserting subtotal rows
// Tags: Aspose.Cells grouped subtotal rows | C# SUBTOTAL formula Aspose.Cells | Excel subtotal insertion Aspose.Cells | dynamic range SUBTOTAL Aspose.Cells | auto‑fit columns after subtotal Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalExample
{
    // The sample creates a workbook, fills columns with category and amount data, detects changes in the category column, inserts a new row for each group's subtotal, writes a label (e.g., "Food Subtotal"), applies a SUBTOTAL(9, Bstart:Bend) formula using the (subtotal1:ColumnName) syntax, updates the start row for the next group, auto‑fits the columns, and saves the file as SubtotalsExample.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data: Column A = Category, Column B = Amount
            // Header
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");

            // Data rows
            string[] categories = { "Food", "Food", "Food", "Transport", "Transport", "Utilities", "Utilities" };
            double[] amounts = { 120.5, 80.0, 45.75, 60.0, 30.25, 100.0, 150.0 };

            for (int i = 0; i < categories.Length; i++)
            {
                cells[i + 2, 0].PutValue(categories[i]); // Column A
                cells[i + 2, 1].PutValue(amounts[i]);   // Column B
            }

            // Determine group boundaries and insert subtotal rows
            int startRow = 2; // First data row (1-based index)
            string currentCategory = categories[0];

            for (int i = 0; i < categories.Length; i++)
            {
                // When category changes or at the last row, insert subtotal
                bool isLastRow = (i == categories.Length - 1);
                bool categoryChanged = !isLastRow && categories[i + 1] != currentCategory;

                if (categoryChanged || isLastRow)
                {
                    int endRow = i + 2; // End row of current group (1-based)

                    // Insert a new row for subtotal
                    sheet.Cells.InsertRows(endRow, 1);
                    // Write label
                    cells[endRow, 0].PutValue($"{currentCategory} Subtotal");

                    // Insert SUBTOTAL formula for the Amount column (Column B)
                    // Syntax: =SUBTOTAL(9, B{startRow}:B{endRow})
                    // 9 = SUM function in SUBTOTAL
                    string formula = $"SUBTOTAL(9,B{startRow}:B{endRow - 1})";
                    cells[endRow, 1].Formula = formula;

                    // Update startRow for next group (skip the inserted subtotal row)
                    startRow = endRow + 1;
                    // Update currentCategory for next group if not last row
                    if (!isLastRow)
                        currentCategory = categories[i + 1];
                }
            }

            // Auto-fit columns for better visibility
            sheet.AutoFitColumns();

            // Save the workbook
            workbook.Save("SubtotalsExample.xlsx");
        }
    }
}
