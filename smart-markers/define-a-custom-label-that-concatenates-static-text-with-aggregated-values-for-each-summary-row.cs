// Title: Create a custom totals row label that combines static text with a SUM formula using Aspose.Cells ListObject in C#
// AI Prompts: Generate C# code with Aspose.Cells that adds a ListObject table, enables a totals row, and sets the totals cell formula to concatenate a fixed label and the SUM of a column. | Show how to build the data range address dynamically and apply it in a concatenated text‑and‑SUM formula for the totals row in an Excel workbook using Aspose.Cells. | Demonstrate assigning a custom label to the first column of the totals row while using a formula like "Grand Total: " & SUM(...) in Aspose.Cells for .NET.
// Common Searches: aspnet how to set a custom label and sum formula in a totals row with Aspose.Cells | c# Aspose.Cells concatenate text with SUM in table totals row | create dynamic data range for totals row formula using Aspose.Cells ListObject | add totals row with custom label in Excel file via Aspose.Cells C# | Aspose.Cells example for custom totals row text and calculation
// Tags: set totals row formula Aspose.Cells | custom totals row labeling ListObject | concatenate static text with SUM C# | dynamic data range address Aspose.Cells | Excel table totals calculation .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace CustomTotalsLabelDemo
{
    // The sample creates an Excel workbook, inserts product and price data, adds a ListObject table with a totals row, configures the Price column to sum values, builds the data range address dynamically, sets the totals cell formula to "Grand Total: " & SUM(range), assigns a "Totals" label to the first column of the totals row, and saves the file as CustomTotalsLabelDemo.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                // Header row
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Price");

                // Data rows
                sheet.Cells["A2"].PutValue("Item1");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["A3"].PutValue("Item2");
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["A4"].PutValue("Item3");
                sheet.Cells["B4"].PutValue(200);

                // Add a table that includes the header, data and a totals row
                // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
                int tableIndex = sheet.ListObjects.Add(0, 0, 4, 1, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.ShowTotals = true; // Enable the totals row

                // Configure the totals calculation for the Price column (index 1)
                ListColumn priceColumn = table.ListColumns[1];
                priceColumn.TotalsCalculation = TotalsCalculation.Sum;

                // Determine the row index of the totals row (zero‑based)
                // Totals row is placed immediately after the last data row of the table
                int totalsRowIndex = table.DataRange.FirstRow + table.DataRange.RowCount;

                // Build the range address of the data cells in the Price column (e.g., B2:B4)
                string firstDataCell = sheet.Cells[table.DataRange.FirstRow + 1, 1].Name; // B2
                string lastDataCell = sheet.Cells[table.DataRange.FirstRow + table.DataRange.RowCount, 1].Name; // B4
                string dataRange = $"{firstDataCell}:{lastDataCell}";

                // Set a formula in the totals cell that concatenates static text with the aggregated sum
                // Excel formula: ="Grand Total: " & SUM(B2:B4)
                sheet.Cells[totalsRowIndex, 1].Formula = $"\"Grand Total: \" & SUM({dataRange})";

                // Optionally set a label for the first column in the totals row
                ListColumn productColumn = table.ListColumns[0];
                productColumn.TotalsRowLabel = "Totals";

                // Save the workbook
                workbook.Save("CustomTotalsLabelDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
