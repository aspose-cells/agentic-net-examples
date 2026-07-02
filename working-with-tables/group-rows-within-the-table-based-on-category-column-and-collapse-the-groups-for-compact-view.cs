using System;
using Aspose.Cells;

namespace AsposeCellsGroupRowsByCategory
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // Sample data: Column A = Category, Column B = Value
            // ------------------------------------------------------------
            // Header
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");

            // Data rows (starting at row index 1, i.e., Excel row 2)
            string[] categories = { "Alpha", "Alpha", "Beta", "Beta", "Beta", "Gamma", "Gamma", "Alpha" };
            int[] values = { 10, 15, 20, 25, 30, 35, 40, 45 };

            for (int i = 0; i < categories.Length; i++)
            {
                int rowIndex = i + 1; // offset for header
                cells[rowIndex, 0].PutValue(categories[i]); // Category column (A)
                cells[rowIndex, 1].PutValue(values[i]);    // Value column (B)
            }

            // ------------------------------------------------------------
            // Group rows based on contiguous identical categories
            // ------------------------------------------------------------
            int firstDataRow = 1; // zero‑based index of the first data row
            int lastDataRow = firstDataRow + categories.Length - 1;

            // Initialize the start of the first group
            int groupStart = firstDataRow;
            string currentCategory = cells[groupStart, 0].StringValue;

            for (int row = firstDataRow + 1; row <= lastDataRow; row++)
            {
                string category = cells[row, 0].StringValue;

                // When the category changes, close the previous group
                if (!category.Equals(currentCategory, StringComparison.Ordinal))
                {
                    // Group rows from groupStart to row‑1 and hide them (collapsed view)
                    cells.GroupRows(groupStart, row - 1, true);

                    // Start a new group
                    groupStart = row;
                    currentCategory = category;
                }
            }

            // Group the final range
            cells.GroupRows(groupStart, lastDataRow, true);

            // Optional: place the summary row above the detail rows for a compact outline
            worksheet.Outline.SummaryRowBelow = false;

            // ------------------------------------------------------------
            // Save the workbook
            // ------------------------------------------------------------
            workbook.Save("GroupedByCategory.xlsx");
        }
    }
}