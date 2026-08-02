// Title: Add a Totals Row with Sum Calculations to All Excel Tables Using Aspose.Cells for .NET
// Description: C# code that loads an Excel workbook, iterates through each worksheet and ListObject, turns on the totals row when it is absent, assigns TotalsCalculation.Sum to every column, and saves the updated file.
// Keywords: Aspose.Cells | .NET | C# | Excel totals row | ListObject | ShowTotals | TotalsCalculation.Sum | add totals row programmatically | detect tables without totals row | Excel table automation
// Common Searches: How to add a totals row to all tables in an Excel workbook with Aspose.Cells | Aspose.Cells C# enable totals row for ListObject | Set default sum totals for Excel tables using Aspose.Cells | Iterate worksheets and add missing totals row Aspose.Cells | Programmatically show totals row in Excel tables .NET
// Developer Intent: Enable a missing totals row and apply a sum calculation to each column of every table in a workbook.
// Use Cases: Automatically insert a totals row in financial reports so each table displays column sums before distribution. | Prepare data‑export files that must contain sum totals for downstream processing. | Standardize Excel templates by guaranteeing every ListObject includes a totals row with default calculations.
// AI Prompts: Generate C# code with Aspose.Cells that adds a totals row using average calculations instead of sum for numeric columns. | Create a method that adds a sum totals row only when a table has no custom totals, preserving existing settings. | Add robust error handling for missing input files and ensure the workbook is saved after all tables are processed.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTotalsRowAdder
{
    // C# code that loads an Excel workbook, iterates through each worksheet and ListObject, turns on the totals row when it is absent, assigns TotalsCalculation.Sum to every column, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Access the collection of tables (ListObjects) on the current worksheet
                ListObjectCollection tables = worksheet.ListObjects;

                // Process each table
                foreach (ListObject table in tables)
                {
                    // If the table does not already show a totals row, enable it
                    if (!table.ShowTotals)
                    {
                        table.ShowTotals = true;

                        // Set a default Sum calculation for each column in the totals row
                        for (int colIndex = 0; colIndex < table.ListColumns.Count; colIndex++)
                        {
                            table.ListColumns[colIndex].TotalsCalculation = TotalsCalculation.Sum;
                        }
                    }
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}
