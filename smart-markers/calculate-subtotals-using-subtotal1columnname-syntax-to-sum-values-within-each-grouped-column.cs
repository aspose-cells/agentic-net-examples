// Title: Calculate grouped subtotals in Excel with Aspose.Cells C# (subtotal1:ColumnName syntax)
// Description: Shows how to use Aspose.Cells Cells.Subtotal to group rows by a column, sum another column, replace existing subtotals, insert page breaks, place summary rows below each group, and save the workbook.
// Keywords: Aspose.Cells | C# subtotal | Excel group subtotal | Cells.Subtotal method | subtotal1 column syntax | smart markers subtotal | Excel report pagination | page breaks Aspose | summary below data | group by column Aspose
// Common Searches: Aspose.Cells subtotal example C# | How to use Cells.Subtotal in C# | Group rows and sum column Aspose.Cells | Add page breaks with subtotal Aspose | subtotal1:ColumnName smart marker | Retrieve SubtotalSetting Aspose.Cells | C# Excel subtotal rows
// Developer Intent: Add grouped sum subtotals, page breaks, and summary rows to an Excel worksheet using Aspose.Cells.
// Use Cases: Create a regional financial report that automatically inserts subtotal rows for each region and adds page breaks for printable sections. | Generate a sales summary where each product category receives a summed subtotal row placed directly below its grouped data. | Export data to Excel with grouped totals and pagination, enabling easy distribution of large datasets across multiple pages.
// AI Prompts: Write C# code with Aspose.Cells to apply Cells.Subtotal, grouping by the first column and summing the second column, with page breaks and summary rows below each group. | Explain how to retrieve and modify SubtotalSetting after calling Cells.Subtotal in Aspose.Cells. | Show how to use the (subtotal1:ColumnName) syntax inside a smart marker template to calculate grouped subtotals automatically.

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalExample
{
    // Shows how to use Aspose.Cells Cells.Subtotal to group rows by a column, sum another column, replace existing subtotals, insert page breaks, place summary rows below each group, and save the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data
            // Header row
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");

            // Data rows
            object[,] data = new object[,]
            {
                { "North", 1200 },
                { "North", 800 },
                { "South", 1500 },
                { "South", 700 },
                { "East",  900 },
                { "East",  1100 }
            };

            for (int i = 0; i < data.GetLength(0); i++)
            {
                cells[i + 1, 0].PutValue(data[i, 0]); // Category column (A)
                cells[i + 1, 1].PutValue(data[i, 1]); // Amount column (B)
            }

            // Define the range that contains the data (including header)
            // A1:B7 -> rows 0‑6, columns 0‑1
            CellArea area = CellArea.CreateCellArea(0, 0, data.GetLength(0), 1);

            // Apply subtotal:
            // - Group by the first column (Category) -> index 0
            // - Use SUM function
            // - Add subtotal for the second column (Amount) -> index 1
            // - Replace existing subtotals, add page breaks, place summary below data
            cells.Subtotal(
                area,
                0,                                 // groupBy column index
                ConsolidationFunction.Sum,         // subtotal function
                new int[] { 1 },                   // columns to subtotal
                true,                              // replace existing subtotals
                true,                              // insert page breaks between groups
                true                               // place summary below each group
            );

            // OPTIONAL: Retrieve the subtotal setting to verify parameters
            SubtotalSetting setting = cells.RetrieveSubtotalSetting(area);
            Console.WriteLine($"GroupBy column index: {setting.GroupBy}");
            Console.WriteLine($"Subtotal function: {setting.SubtotalFunction}");
            Console.WriteLine($"Subtotal column index: {setting.TotalList[0]}");
            Console.WriteLine($"Summary below data: {setting.SummaryBelowData}");

            // Save the workbook
            workbook.Save("SubtotalExample.xlsx");
        }
    }
}
