// Title: Group rows and auto‑fit heights with Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, add sample data, group a range of rows (e.g., rows 2‑5) without collapsing them, auto‑fit the grouped rows to their content, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | group rows | auto fit rows | Worksheet.AutoFitRows | Cells.GroupRows | Excel row grouping programmatically | adjust row height | collapse rows | Excel automation
// Common Searches: Aspose.Cells group rows C# | auto fit grouped rows Aspose.Cells | how to group rows without hiding in Aspose.Cells | C# code to auto‑fit row height after grouping | Excel row grouping using Aspose.Cells .NET
// Developer Intent: Programmatically group a set of rows and automatically adjust their height.
// Use Cases: Generate financial reports where related line items are grouped and fully visible. | Export hierarchical data to Excel with collapsible sections that retain proper row height. | Create printable worksheets with grouped rows that auto‑adjust to the content.
// AI Prompts: Write C# code using Aspose.Cells to group rows 10‑20 and auto‑fit them. | Provide a reusable method that takes startRow and endRow, groups the rows without hiding, and calls AutoFitRows on the same range. | Explain how to conditionally group rows based on cell values and ensure the grouped rows are auto‑fit for proper display. | Show how to programmatically expand or collapse grouped rows after auto‑fitting them.

using System;
using Aspose.Cells;

namespace AsposeCellsRowGroupingDemo
{
    // Shows how to create a workbook, add sample data, group a range of rows (e.g., rows 2‑5) without collapsing them, auto‑fit the grouped rows to their content, and save the file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data (optional, helps visualize the grouping)
            for (int i = 0; i < 10; i++)
            {
                cells[i, 0].PutValue($"Row {i + 1}");
                cells[i, 1].PutValue($"Data {i + 1}");
            }

            // Group rows 2 through 5 (zero‑based indices 1 to 4) without hiding them
            // Uses Cells.GroupRows(int firstIndex, int lastIndex, bool isHidden)
            cells.GroupRows(1, 4, false);

            // Auto‑fit the grouped rows to adjust their heights based on content
            // Uses Worksheet.AutoFitRows(int startRow, int endRow)
            worksheet.AutoFitRows(1, 4);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("GroupedRowsAutoFitDemo.xlsx");
        }
    }
}
