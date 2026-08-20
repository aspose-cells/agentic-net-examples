// Title: Add Missing Totals Row with Sum Calculations to Excel Tables using Aspose.Cells for .NET
// Description: Loads a workbook, scans every worksheet for ListObjects without a totals row, enables the totals row, applies a SUM calculation to each column, optionally labels the first cell, and saves the updated file.
// Keywords: Aspose.Cells totals row | C# ListObject ShowTotals | Excel table sum calculation | programmatic totals row | add missing totals row | Aspose.Cells .NET example | ListObject TotalsCalculation.Sum | Excel automation C#
// Common Searches: How to enable a totals row for Excel tables with Aspose.Cells C# | Programmatically add sum totals to ListObjects in a workbook | Detect tables without totals row and insert one using Aspose.Cells | C# code to set TotalsCalculation.Sum for all columns | Aspose.Cells example for adding a totals row automatically
// Developer Intent: Insert a default SUM totals row into every Excel table that lacks one.
// Use Cases: Generate consolidated summary rows for financial statements across multiple sheets. | Standardize reporting templates before distributing workbooks to clients. | Automate data‑import pipelines that create tables without totals, ensuring each table displays column totals. | Prepare analytics dashboards where every dataset needs a quick aggregate row.
// AI Prompts: Create a reusable C# method with Aspose.Cells that adds a totals row using average calculations instead of sum. | Write code to format the inserted totals row with bold text, a gray background, and right‑aligned numbers. | Generate a script that adds a custom label to the first column's totals cell and applies currency formatting to numeric totals.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTotalsRowAdder
{
    // Loads a workbook, scans every worksheet for ListObjects without a totals row, enables the totals row, applies a SUM calculation to each column, optionally labels the first cell, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Iterate through all tables (ListObjects) in the current worksheet
                foreach (ListObject table in worksheet.ListObjects)
                {
                    // If the table does not already show a totals row, add one
                    if (!table.ShowTotals)
                    {
                        // Enable the totals row for the table
                        table.ShowTotals = true;

                        // Set default sum calculation for each column in the totals row
                        for (int i = 0; i < table.ListColumns.Count; i++)
                        {
                            ListColumn column = table.ListColumns[i];
                            column.TotalsCalculation = TotalsCalculation.Sum;

                            // Optionally, set a label for the first column's totals cell
                            if (i == 0)
                            {
                                column.TotalsRowLabel = "Total";
                            }
                        }
                    }
                }
            }

            // Save the modified workbook (replace with desired output path)
            workbook.Save("output.xlsx");
        }
    }
}
