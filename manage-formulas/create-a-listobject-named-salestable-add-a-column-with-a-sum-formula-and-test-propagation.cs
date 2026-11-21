// Title: Create a ListObject named SalesTable, add a calculated SUM column, and verify formula propagation with Aspose.Cells for .NET (C#)
// Common Searches: how to add a calculated column to a ListObject using Aspose.Cells C# | extend an Aspose.Cells table and keep formulas auto‑filled | resize a ListObject in Aspose.Cells to include new rows with formulas | set row‑level SUM formula for a column in an Aspose.Cells table | assign a display name to a ListObject and save the workbook with Aspose.Cells
// Tags: listobject add calculated column Aspose.Cells | table resize include new rows Aspose.Cells | auto‑propagate formulas in Aspose.Cells tables | row‑wise sum formula Aspose.Cells C# | listobject display name Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;

// The example creates a workbook, defines a ListObject called SalesTable, adds a 'SumQtyPrice' column with a row‑specific SUM formula (Quantity + Price), resizes the table to include the new column, appends an extra data row, resizes the table again to incorporate the new row, confirms that the formula auto‑fills for the added row, and saves the file as SalesTable.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data (header + 3 rows)
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");

            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(2);

            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(5);
            sheet.Cells["C3"].PutValue(1.5);

            sheet.Cells["A4"].PutValue("Orange");
            sheet.Cells["B4"].PutValue(8);
            sheet.Cells["C4"].PutValue(2.2);

            // Define the range for the ListObject (including header)
            int firstRow = 0;          // Row 1 (zero‑based)
            int firstColumn = 0;       // Column A
            int lastRow = 3;           // Row 4 (header + 3 data rows)
            int lastColumn = 2;        // Column C

            // Add a ListObject (table) with headers and give it a display name
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, lastRow, lastColumn, true);
            ListObject salesTable = sheet.ListObjects[tableIndex];
            salesTable.DisplayName = "SalesTable";

            // Add a new column that will contain a SUM formula for each row (Quantity + Price)
            int sumColumnIndex = lastColumn + 1; // Next column after existing ones
            sheet.Cells[firstRow, sumColumnIndex].PutValue("SumQtyPrice");

            // Extend the table to include the new column (hasHeaders = true)
            salesTable.Resize(firstRow, firstColumn, lastRow, sumColumnIndex, true);

            // Set the SUM formula for each data row in the new column
            for (int r = firstRow + 1; r <= lastRow; r++) // skip header row
            {
                // Formula: =B{row}+C{row}
                string formula = $"=B{r + 1}+C{r + 1}";
                sheet.Cells[r, sumColumnIndex].Formula = formula;
            }

            // Test propagation: add a new data row and verify the formula is auto‑filled
            int newRow = lastRow + 1; // Next row index (zero‑based)
            sheet.Cells[newRow, firstColumn].PutValue("Grapes");
            sheet.Cells[newRow, firstColumn + 1].PutValue(12);
            sheet.Cells[newRow, firstColumn + 2].PutValue(3);

            // Resize the table to include the new row (hasHeaders = true)
            salesTable.Resize(firstRow, firstColumn, newRow, sumColumnIndex, true);

            // The formula should have propagated to the new row automatically.
            // Retrieve the calculated value to confirm (12 + 3 = 15)
            double calculatedValue = sheet.Cells[newRow, sumColumnIndex].DoubleValue;
            Console.WriteLine($"Calculated SumQtyPrice for new row: {calculatedValue}");

            // Save the workbook
            workbook.Save("SalesTable.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
