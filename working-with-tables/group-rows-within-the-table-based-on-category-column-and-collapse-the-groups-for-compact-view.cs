// Title: Group and Collapse Rows by Category with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills column A with category labels and column B with data, then automatically groups consecutive rows that share the same category, collapses the groups, places the summary row above the details, and saves the file as GroupedByCategory.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells group rows C# | collapse rows Aspose.Cells | Excel outline groups Aspose.Cells | group rows by column value C# | hide rows programmatically Aspose.Cells | summary row above detail Aspose.Cells | working with tables Aspose.Cells | GitHub Aspose.Cells example | .NET Excel grouping rows
// Common Searches: How to group rows by category in Excel using Aspose.Cells for .NET | Collapse grouped rows with Aspose.Cells C# | Set summary row above detail rows in Aspose.Cells outline | Create outline groups automatically Aspose.Cells | Aspose.Cells example for grouping rows by column
// Developer Intent: Automatically detect consecutive rows with identical category values, create outline groups for them, collapse the groups for a compact view, and configure the summary row position.
// Use Cases: Financial statements where transactions are grouped by account type and initially hidden for a cleaner report. | Data‑analysis workbooks that let users expand or collapse product categories on demand. | Printable inventory lists with category headings that can be collapsed to reduce page length.
// AI Prompts: Generate C# code using Aspose.Cells to group rows based on repeated values in a column and hide the groups. | Explain how to set the outline's SummaryRowBelow property so the summary row appears above the grouped rows. | Provide a reusable method that accepts a worksheet and a column index, then creates collapsed groups for each distinct value.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills column A with category labels and column B with data, then automatically groups consecutive rows that share the same category, collapses the groups, places the summary row above the details, and saves the file as GroupedByCategory.xlsx using Aspose.Cells for .NET.
    public class GroupRowsByCategoryDemo
    {
        // Entry point for the console application
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data: Column A = Category, Column B = Value (rows 0..9 correspond to Excel rows 1..10)
            string[] categories = { "A", "A", "A", "B", "B", "C", "C", "C", "C", "D" };
            for (int i = 0; i < categories.Length; i++)
            {
                cells[i, 0].PutValue(categories[i]);          // Category
                cells[i, 1].PutValue($"Data {i + 1}");       // Some other data
            }

            // Group rows that share the same category (data assumed to be sorted by category)
            int groupStart = 0;
            for (int row = 1; row <= categories.Length; row++)
            {
                // Determine if we have reached the end of the current category block
                bool endOfBlock = row == categories.Length || categories[row] != categories[row - 1];
                if (endOfBlock)
                {
                    int groupEnd = row - 1; // inclusive index of the last row in the block
                    if (groupEnd > groupStart) // Only group if there is more than one row
                    {
                        // Collapse (hide) the grouped rows for a compact view
                        cells.GroupRows(groupStart, groupEnd, true);
                    }
                    // Start a new group at the current row
                    groupStart = row;
                }
            }

            // Place the summary row above the detail rows (default is true)
            worksheet.Outline.SummaryRowBelow = false;

            // Save the workbook
            workbook.Save("GroupedByCategory.xlsx");
        }
    }
}
