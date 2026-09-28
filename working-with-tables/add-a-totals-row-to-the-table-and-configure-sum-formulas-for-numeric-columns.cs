// Title: Add a totals row with SUM formulas to an Aspose.Cells ListObject table using C#
// AI Prompts: Create a ListObject from a data range, enable its totals row, and assign SUM formulas to numeric columns with Aspose.Cells for .NET. | Update an existing Aspose.Cells table to display a totals row that automatically sums selected columns. | Insert a custom label in the totals row and apply a currency format to the summed Price column using the Aspose.Cells C# API.
// Common Searches: aspnet add totals row to Excel table using Aspose.Cells ListObject | c# set sum formula for table column in Aspose.Cells workbook | how to enable totals row in Aspose.Cells ListObject programmatically | Aspose.Cells calculate column totals with SUM formula in C#
// Tags: Aspose.Cells ListObject totals row | C# set SUM formula in Excel table | Aspose.Cells add table totals row | Excel table sum column Aspose.Cells | Aspose.Cells generate totals row programmatically

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a new workbook, fills it with sample data, converts the range into a ListObject named "SalesData", shows a totals row, places a "Total" label, and assigns SUM formulas to the Quantity and Price columns before saving the file as TotalsTable.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (headers + numeric columns)
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");

            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(0.5);

            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["C3"].PutValue(0.3);

            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(15);
            sheet.Cells["C4"].PutValue(0.8);

            // Define the range that will become a table (including header row)
            int startRow = 0;      // zero‑based index for row 1
            int startColumn = 0;   // zero‑based index for column A
            int endRow = 4;        // last data row (row 5 in Excel)
            int endColumn = 2;     // column C

            // Add a ListObject (table) to the worksheet
            int totalRows = endRow - startRow + 1;
            int totalColumns = endColumn - startColumn + 1;
            int tableIndex = sheet.ListObjects.Add(startRow, startColumn, totalRows, totalColumns, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Set a display name for the table
            table.DisplayName = "SalesData";

            // Show the totals row
            table.ShowTotals = true;

            // Row index of the totals row (zero‑based)
            int totalsRowIndex = endRow + 1;

            // Set label for the first column in the totals row
            sheet.Cells[totalsRowIndex, startColumn].PutValue("Total");

            // Configure SUM formulas for numeric columns (Quantity and Price)
            // Column indexes are zero‑based relative to the worksheet
            for (int col = 1; col <= 2; col++)
            {
                // Convert column index to Excel column name (e.g., 1 -> "B")
                string colName = CellsHelper.ColumnIndexToName(startColumn + col);

                // Build the SUM formula covering data rows (excluding header and totals row)
                // Data starts at row 2 (Excel row 2) and ends at row endRow + 1 (Excel row 5)
                string formula = $"SUM({colName}2:{colName}{endRow + 1})";

                // Assign the formula to the corresponding column in the totals row
                sheet.Cells[totalsRowIndex, startColumn + col].Formula = formula;
            }

            // Save the workbook
            workbook.Save("TotalsTable.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
