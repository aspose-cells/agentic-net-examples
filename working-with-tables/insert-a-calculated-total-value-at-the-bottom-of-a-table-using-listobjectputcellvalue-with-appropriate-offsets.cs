// Title: Add a Total Row to an Aspose.Cells ListObject Table with PutCellValue (C#)
// Description: Demonstrates how to compute the sum of a numeric column, calculate the correct row offset, and insert a "Total" label and value at the bottom of a ListObject table using Aspose.Cells' ListObject.PutCellValue method. The example creates a workbook, builds a simple table, iterates the data, and saves the result as an XLSX file.
// Keywords: Aspose.Cells ListObject PutCellValue | C# add total row to table | Aspose.Cells calculate column sum | insert summary row Aspose.Cells | .NET spreadsheet total footer | dynamic table offset Aspose.Cells
// Common Searches: how to add a total row to a ListObject in Aspose.Cells | PutCellValue offset example C# | sum column and write total in Aspose.Cells table | append footer row to Aspose.Cells ListObject | Aspose.Cells table summary row code
// Developer Intent: Insert a calculated total row at the end of a ListObject table using ListObject.PutCellValue with the proper relative offset.
// Use Cases: Generate invoices where the grand total appears automatically below the items table. | Create sales or expense reports that add a dynamic total row without hard‑coding cell addresses. | Build reusable spreadsheet utilities that summarize any numeric column in a ListObject.
// AI Prompts: Write C# code with Aspose.Cells to add a total row to an existing ListObject, compute the sum of a numeric column, and place the label and sum using PutCellValue and a relative row offset. | Show an Aspose.Cells .NET example that iterates over a ListObject to calculate a column total and appends a summary row with PutCellValue, deriving the row index from the table's start and end rows. | Explain how to determine the correct offset for inserting a footer row after a ListObject and how to use ListObject.PutCellValue to write values into that row.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to compute the sum of a numeric column, calculate the correct row offset, and insert a "Total" label and value at the bottom of a ListObject table using Aspose.Cells' ListObject.PutCellValue method. The example creates a workbook, builds a simple table, iterates the data, and saves the result as an XLSX file.
class InsertTotalWithPutCellValue
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (header + 3 rows)
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["B1"].PutValue("Amount");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(80);
        sheet.Cells["A4"].PutValue("Cherry");
        sheet.Cells["B4"].PutValue(150);

        // Create a ListObject (table) that includes the data range
        int tableIndex = sheet.ListObjects.Add("A1", "B4", true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Compute the sum of the Amount column (column index 1)
        double sum = 0;
        for (int row = table.StartRow + 1; row <= table.EndRow; row++) // data rows only
        {
            object val = sheet.Cells[row, table.StartColumn + 1].Value;
            if (val is double d) sum += d;
            else if (val is int i) sum += i;
        }

        // Determine the offset for the new total row (one row after the last data row)
        int totalRowOffset = table.EndRow - table.StartRow + 1; // relative offset

        // Insert a label in the first column of the total row
        table.PutCellValue(totalRowOffset, 0, "Total");

        // Insert the calculated sum in the Amount column of the total row
        table.PutCellValue(totalRowOffset, 1, sum);

        // Save the workbook
        workbook.Save("TableWithTotal.xlsx", SaveFormat.Xlsx);
    }
}
