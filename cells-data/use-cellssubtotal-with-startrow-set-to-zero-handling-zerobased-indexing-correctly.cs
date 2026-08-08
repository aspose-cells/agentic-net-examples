// Title: Apply Cells.Subtotal with Zero‑Based StartRow in Aspose.Cells (C#)
// Description: Creates a workbook, adds a header and sales rows, defines a CellArea that starts at row 0, and uses Cells.Subtotal to group by the first column (Region) and sum the Sales column. The example also shows how to retrieve the SubtotalSetting before saving the file as SubtotalDemo.xlsx.
// Keywords: Aspose.Cells | Cells.Subtotal | zero based indexing | C# | CellArea | group by column | sum subtotal | SubtotalSetting | Excel report generation | US developers | India developers
// Common Searches: Aspose.Cells Cells.Subtotal startRow 0 C# example | how to use zero‑based indexing with Cells.Subtotal | group rows by column and sum values Aspose.Cells | retrieve SubtotalSetting after applying Cells.Subtotal | define CellArea that includes header row Aspose.Cells
// Developer Intent: Add subtotal rows to an Excel worksheet using Aspose.Cells while correctly handling zero‑based row indices.
// Use Cases: Create a regional sales summary that automatically inserts subtotal rows for each region. | Validate subtotal configuration programmatically by reading the SubtotalSetting object. | Export a pre‑formatted report with grouped totals for downstream analysis or sharing.
// AI Prompts: Generate C# code to place subtotal rows above the data instead of below using Aspose.Cells. | Show how to apply multiple subtotal columns (e.g., Sales and Quantity) with different functions in one call. | Explain how to hide or collapse subtotal rows after they are created with Cells.Subtotal.

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalDemo
{
    // Creates a workbook, adds a header and sales rows, defines a CellArea that starts at row 0, and uses Cells.Subtotal to group by the first column (Region) and sum the Sales column. The example also shows how to retrieve the SubtotalSetting before saving the file as SubtotalDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (including header row at index 0)
            // Header row
            cells[0, 0].PutValue("Region");   // Column A
            cells[0, 1].PutValue("Product");  // Column B
            cells[0, 2].PutValue("Sales");    // Column C

            // Data rows start at row index 1 (zero‑based)
            object[,] data = new object[,]
            {
                { "North", "Widget", 5000 },
                { "North", "Gadget", 3000 },
                { "South", "Widget", 6000 },
                { "South", "Gadget", 4000 },
                { "West",  "Widget", 4500 }
            };

            for (int r = 0; r < data.GetLength(0); r++)
            {
                for (int c = 0; c < data.GetLength(1); c++)
                {
                    // Row index = r + 1 because data starts after header
                    cells[r + 1, c].PutValue(data[r, c]);
                }
            }

            // Define the range that includes the header row (A1:C6)
            // StartRow = 0 ensures zero‑based indexing is handled correctly
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = data.GetLength(0),   // last data row index (5) because header + 5 data rows = 6 rows total, zero‑based end row = 5
                EndColumn = 2
            };

            // Apply subtotal:
            // - Group by the first column (Region) -> groupBy = 0 (zero‑based)
            // - Use SUM function
            // - Add subtotal to the third column (Sales) -> totalList = new int[] { 2 } (zero‑based)
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 2 });

            // Optional: retrieve and display subtotal settings
            SubtotalSetting setting = cells.RetrieveSubtotalSetting(area);
            Console.WriteLine($"GroupBy index: {setting.GroupBy}");
            Console.WriteLine($"Subtotal function: {setting.SubtotalFunction}");
            Console.WriteLine($"Total columns: {string.Join(",", setting.TotalList)}");
            Console.WriteLine($"Summary below data: {setting.SummaryBelowData}");

            // Save the workbook (save rule)
            workbook.Save("SubtotalDemo.xlsx");
        }
    }
}
