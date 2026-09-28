// Title: Create StdDev subtotals for column H and place the summary row at the bottom with Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells that groups rows by column A, adds a standard deviation subtotal for column H, and positions the summary row below the data. | Show how to call Cells.Subtotal with ConsolidationFunction.StdDev, enable page breaks, and set summaryBelowData to true in a .NET workbook. | Provide a complete example that creates sample data, applies StdDev subtotals on column H, and saves the workbook as SubtotalStdDevBottom.xlsx.
// Common Searches: Aspose.Cells C# add standard deviation subtotal to a column and put summary at the bottom | How to use Cells.Subtotal with StdDev and page breaks in .NET | Group rows by first column and calculate StdDev subtotal for column H using Aspose.Cells
// Tags: Aspose.Cells Cells.Subtotal StdDev | C# group rows by column A subtotal | Excel standard deviation subtotal Aspose | summary row bottom placement Aspose.Cells | add page breaks Aspose.Cells subtotal

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalStdDevDemo
{
    // Demonstrates creating a workbook, populating sample data, and using Cells.Subtotal to group by column A, calculate standard deviation for column H, add page breaks, place the summary row at the bottom, and save the file as SubtotalStdDevBottom.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // Sample data creation (optional, replace with your own data)
            // ------------------------------------------------------------
            // Header row
            cells["A1"].PutValue("Group");
            cells["H1"].PutValue("Values");

            // Populate some sample numeric data in column H (index 7)
            for (int row = 1; row <= 10; row++)
            {
                // Example group values in column A (index 0)
                cells[row, 0].PutValue(row % 2 == 0 ? "Even" : "Odd");
                // Random numbers in column H
                cells[row, 7].PutValue(10 + row * 5);
            }

            // ------------------------------------------------------------
            // Define the range that contains the data for subtotals
            // Here we assume data starts at A1 and ends at H10 (adjust as needed)
            // ------------------------------------------------------------
            CellArea area = CellArea.CreateCellArea("A1", "H10");

            // ------------------------------------------------------------
            // Add subtotals:
            // - Group by the first column (index 0, "Group")
            // - Use StdDev function (ConsolidationFunction.StdDev)
            // - Apply subtotal to column H (index 7)
            // - Replace existing subtotals, add page breaks, place summary below data (bottom)
            // ------------------------------------------------------------
            cells.Subtotal(
                area,                     // range
                0,                        // groupBy column index (A)
                ConsolidationFunction.StdDev, // StdDev function
                new int[] { 7 },          // subtotal on column H
                true,                     // replace existing subtotals
                true,                     // add page breaks between groups
                true);                    // summaryBelowData = true (bottom)

            // ------------------------------------------------------------
            // Save the workbook
            // ------------------------------------------------------------
            workbook.Save("SubtotalStdDevBottom.xlsx");
        }
    }
}
