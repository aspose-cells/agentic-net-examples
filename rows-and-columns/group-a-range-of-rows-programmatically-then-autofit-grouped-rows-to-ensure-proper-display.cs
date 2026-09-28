// Title: How to programmatically group rows and auto‑fit their heights in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to group rows 2‑6, hide the group, and call AutoFitRows to adjust the row heights before saving. | Create a workbook, populate sample data, set Outline.SummaryRowBelow = true, then auto‑fit a hidden row range with Aspose.Cells. | Generate an Excel file with grouped rows and automatically calculate appropriate row heights using the AutoFitRows method in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# example for grouping rows and auto fitting row height | auto fit hidden rows after grouping with Aspose.Cells .NET | set outline summary row below grouped rows using Aspose.Cells | how to adjust row height for grouped rows in Excel via Aspose.Cells C# | programmatically group rows and auto‑fit rows in Aspose.Cells workbook
// Tags: group rows Aspose.Cells C# | AutoFitRows hidden rows Aspose.Cells | outline summary row below Aspose.Cells | Excel row grouping programmatically Aspose | adjust row height after grouping Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsGroupRowsAutoFit
{
    // The sample creates a new workbook, fills ten rows with data, groups rows 2‑6 (indices 1‑5) and hides them, sets the outline to place the summary row below the group, auto‑fits the grouped rows to correct their heights, and saves the file as GroupedRowsAutoFitDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (10 rows, 3 columns)
            for (int row = 0; row < 10; row++)
            {
                cells[row, 0].PutValue($"Item {row + 1}");
                cells[row, 1].PutValue($"Description for item {row + 1}");
                cells[row, 2].PutValue($"Details with a relatively long text that may require row height adjustment for row {row + 1}");
            }

            // Group rows 2 through 6 (zero‑based indices 1‑5) and hide them initially
            cells.GroupRows(1, 5, true);

            // Ensure the summary row appears below the grouped rows (optional)
            worksheet.Outline.SummaryRowBelow = true;

            // Auto‑fit the rows that belong to the group to adjust their heights correctly
            // This call will consider the hidden rows as well when calculating the required height
            worksheet.AutoFitRows(1, 5);

            // Save the workbook
            workbook.Save("GroupedRowsAutoFitDemo.xlsx");
        }
    }
}
