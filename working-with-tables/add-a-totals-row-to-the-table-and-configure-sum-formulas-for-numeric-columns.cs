// Title: Add a Totals Row with Sum Formulas to an Aspose.Cells ListObject Table (C#)
// Description: Creates a new workbook, defines a ListObject table with header and data rows, enables the totals row, applies TotalsCalculation.Sum to the Quantity and Price columns, sets a "Grand Total" label for the first column, and saves the file as TotalsRowDemo.xlsx.
// Keywords: Aspose.Cells totals row | C# ListObject sum | Excel table aggregate values | Aspose.Cells Table TotalsCalculation | Add Grand Total label
// Common Searches: Aspose.Cells enable totals row C# | Set sum calculation for table columns Aspose.Cells | How to add Grand Total label in Aspose.Cells ListObject | Create Excel table with automatic totals using Aspose.Cells .NET
// Developer Intent: Insert a totals row into a ListObject table and configure sum calculations for numeric columns.
// Use Cases: Generate invoices where Quantity and Price are summed automatically. | Build sales dashboards that display a Grand Total row without manual formulas. | Export data sets to Excel with a formatted table that shows aggregate totals for numeric fields.
// AI Prompts: Write C# code with Aspose.Cells to add a totals row to a ListObject and set sum calculations for selected columns. | Show how to assign a custom label to the totals row and format the summed values in an Aspose.Cells workbook. | Provide an example of creating a dynamic table range and automatically adding sum formulas to numeric columns using Aspose.Cells .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTotalsRowDemo
{
    // Creates a new workbook, defines a ListObject table with header and data rows, enables the totals row, applies TotalsCalculation.Sum to the Quantity and Price columns, sets a "Grand Total" label for the first column, and saves the file as TotalsRowDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate header row
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Quantity");
            cells["C1"].PutValue("Price");

            // Populate sample data (numeric columns are Quantity and Price)
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(10);
            cells["C2"].PutValue(0.5);

            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(20);
            cells["C3"].PutValue(0.3);

            cells["A4"].PutValue("Cherry");
            cells["B4"].PutValue(15);
            cells["C4"].PutValue(0.8);

            // Define the range of the table (including header and data)
            int firstRow = 0;   // zero‑based index for row 1
            int firstCol = 0;   // column A
            int lastRow = 4;    // row 5 (header + 4 data rows)
            int lastCol = 2;    // column C

            // Add a ListObject (table) to the worksheet (lifecycle rule: create)
            int tableIndex = worksheet.ListObjects.Add(firstRow, firstCol, lastRow, lastCol, true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Enable the totals row (rule: ListObject.ShowTotals)
            table.ShowTotals = true;

            // Set TotalsCalculation = Sum for numeric columns (Quantity and Price)
            // Column indexes in ListObject are zero‑based relative to the table
            table.ListColumns[1].TotalsCalculation = TotalsCalculation.Sum; // Quantity column
            table.ListColumns[2].TotalsCalculation = TotalsCalculation.Sum; // Price column

            // Optionally set a label for the first column in the totals row
            table.ListColumns[0].TotalsRowLabel = "Grand Total";

            // Save the workbook (lifecycle rule: save)
            workbook.Save("TotalsRowDemo.xlsx");
        }
    }
}
