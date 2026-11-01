// Title: Insert subtotal rows for each category in an Excel worksheet using Aspose.Cells with zero‑based row indexing (C#)
// AI Prompts: Generate C# code that groups rows by a column, computes the sum for each group, and inserts a subtotal row using Aspose.Cells while accounting for zero‑based row numbers. | Show how to adjust row offsets when inserting new rows for subtotals in an Aspose.Cells workbook that starts at index 0.
// Common Searches: how to add subtotal rows after each category group using Aspose.Cells in C# with zero based indexing | Aspose.Cells C# insert rows for group totals while handling row offset | calculate category totals and insert subtotal rows in Excel workbook using Aspose.Cells .NET | zero based row indexing Aspose.Cells subtotal example C#
// Tags: Aspose.Cells insert rows | Aspose.Cells calculate subtotals | zero-based row indexing Aspose.Cells | grouped data totals C# Aspose.Cells | Excel subtotal implementation Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example creates a workbook, writes a header and sample category/amount data starting at row 0, computes sums per category with a dictionary, inserts a subtotal row after each category block while adjusting for row offsets, and saves the file as SubtotalResult.xlsx.
class SubtotalExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data starting from row 0 (zero‑based)
            // Header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");

            // Data rows
            string[] categories = { "Food", "Food", "Transport", "Transport", "Utilities", "Utilities" };
            double[] amounts = { 120.5, 80.0, 45.0, 60.0, 100.0, 150.0 };

            for (int i = 0; i < categories.Length; i++)
            {
                // Row index is i + 1 because row 0 holds the header
                sheet.Cells[i + 1, 0].PutValue(categories[i]);   // Column A (index 0)
                sheet.Cells[i + 1, 1].PutValue(amounts[i]);     // Column B (index 1)
            }

            // ----- Manual subtotal implementation -----
            // Calculate sum per category
            var sums = new Dictionary<string, double>();
            for (int i = 0; i < categories.Length; i++)
            {
                if (!sums.ContainsKey(categories[i]))
                    sums[categories[i]] = 0.0;
                sums[categories[i]] += amounts[i];
            }

            // Insert subtotal rows after each category block
            // Keep track of the current insertion offset due to added rows
            int offset = 0;
            string previousCategory = null;
            for (int i = 0; i < categories.Length; i++)
            {
                string currentCategory = categories[i];
                // When category changes (or at the end of the list), insert subtotal for the previous category
                bool isLastRow = (i == categories.Length - 1);
                if (previousCategory != null && (currentCategory != previousCategory || isLastRow))
                {
                    // Determine the row index where the subtotal should be inserted
                    int insertRow = i + 1 + offset; // +1 to place below the last data row of the group
                    // Insert a new row for the subtotal
                    sheet.Cells.InsertRows(insertRow, 1);
                    // Write subtotal label and value
                    sheet.Cells[insertRow, 0].PutValue($"{previousCategory} Total");
                    sheet.Cells[insertRow, 1].PutValue(sums[previousCategory]);
                    // Update offset because we added a row
                    offset++;
                }
                previousCategory = currentCategory;
            }

            // Save the workbook (save rule)
            string outputPath = "SubtotalResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
