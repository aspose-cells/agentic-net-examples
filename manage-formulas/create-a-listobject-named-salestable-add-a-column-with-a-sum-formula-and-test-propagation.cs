// Title: Create a ListObject named SalesTable, add a Total column with SUM formulas, and dynamically resize the table for new rows using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that defines a ListObject, inserts a new column called "Total" with a row‑wise SUM formula, and expands the table to include the column. | Show how to append a data row to an existing Aspose.Cells ListObject, apply the same SUM formula to the new Total cell, and resize the table to incorporate the added row.
// Common Searches: Aspose.Cells add calculated Total column to a ListObject in C# | Resize Aspose.Cells ListObject after inserting rows with formulas | C# set SUM formula for each row in an Excel table using Aspose.Cells | How to expand an Aspose.Cells table to include a newly added column | Append rows to a ListObject and propagate formulas with Aspose.Cells .NET
// Tags: Aspose.Cells ListObject add SUM column | C# resize Aspose.Cells table for new rows | propagate formulas in Aspose.Cells ListObject | dynamic total column generation Aspose.Cells | expand Excel table with Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a workbook, defines a ListObject named SalesTable, inserts a 'Total' column with a SUM formula for each existing row, resizes the table to include the new column, adds an additional data row, applies the same SUM formula to the new Total cell, expands the ListObject to cover the new row, and saves the file as SalesTable.xlsx.
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

            // Populate sample data (header + 4 rows)
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Q1");
            sheet.Cells["C1"].PutValue("Q2");
            sheet.Cells["D1"].PutValue("Q3");

            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["A5"].PutValue("D");

            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);
            sheet.Cells["C5"].PutValue(45);

            sheet.Cells["D2"].PutValue(20);
            sheet.Cells["D3"].PutValue(30);
            sheet.Cells["D4"].PutValue(40);
            sheet.Cells["D5"].PutValue(50);

            // Define the range for the ListObject (table) – header row + 4 data rows, 4 columns
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int lastRow = firstRow + 4;    // includes header (row 0) to row 4
            int lastColumn = firstColumn + 3; // columns A‑D

            // Add the ListObject and name it SalesTable
            int listIndex = sheet.ListObjects.Add(firstRow, firstColumn, lastRow, lastColumn, true);
            ListObject salesTable = sheet.ListObjects[listIndex];
            salesTable.DisplayName = "SalesTable";

            // Insert a new column for Total (SUM) after the existing columns
            int totalColumnIndex = lastColumn + 1; // column E (zero‑based)
            sheet.Cells[firstRow, totalColumnIndex].PutValue("Total");

            // Set SUM formula for each existing data row
            for (int r = firstRow + 1; r <= lastRow; r++) // rows 2‑5 (zero‑based 1‑4)
            {
                // Formula: =SUM(B2:D2) etc.
                string startCell = sheet.Cells[r, firstColumn + 1].Name; // B column
                string endCell = sheet.Cells[r, lastColumn].Name;        // D column
                sheet.Cells[r, totalColumnIndex].Formula = $"=SUM({startCell}:{endCell})";
            }

            // Expand the ListObject to include the new Total column (hasHeaders = true)
            salesTable.Resize(
                salesTable.StartRow,
                salesTable.StartColumn,
                salesTable.DataRange.RowCount,
                salesTable.DataRange.ColumnCount + 1,
                true);

            // ----- Test propagation -----
            // Add a new data row below the current table
            int newRow = salesTable.StartRow + salesTable.DataRange.RowCount; // first empty row after the table
            sheet.Cells[newRow, firstColumn].PutValue("E"); // Product name
            sheet.Cells[newRow, firstColumn + 1].PutValue(12); // Q1
            sheet.Cells[newRow, firstColumn + 2].PutValue(22); // Q2
            sheet.Cells[newRow, firstColumn + 3].PutValue(32); // Q3

            // Apply the same SUM formula to the Total cell of the new row
            string startNew = sheet.Cells[newRow, firstColumn + 1].Name; // B column of new row
            string endNew = sheet.Cells[newRow, lastColumn].Name;        // D column of new row
            sheet.Cells[newRow, totalColumnIndex].Formula = $"=SUM({startNew}:{endNew})";

            // Resize the ListObject to include the newly added row (hasHeaders = true)
            salesTable.Resize(
                salesTable.StartRow,
                salesTable.StartColumn,
                salesTable.DataRange.RowCount + 1,
                salesTable.DataRange.ColumnCount,
                true);

            // Save the workbook
            workbook.Save("SalesTable.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
