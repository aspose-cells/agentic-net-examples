// Title: Add a calculated total row to the bottom of an Excel ListObject using ListObject.PutCellValue in Aspose.Cells for .NET (C#)
// Common Searches: how to insert a sum row at the end of a ListObject table with Aspose.Cells C# | using ListObject.PutCellValue to write a total value in an Excel table via Aspose.Cells | calculate column total and add a footer row in a .NET Excel workbook with Aspose.Cells | C# code to add a summary row to an Excel table created by Aspose.Cells | Aspose.Cells ListObject add total row programmatically in C#
// Tags: listobject putcellvalue add total row | aspose.cells calculate column sum c# | excel table footer insertion aspnet | c# insert summary row into listobject | aspose.cells datarange total calculation

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;

// The example creates a workbook, defines a ListObject covering A1:B5, computes the sum of the 'Amount' column, and uses ListObject.PutCellValue with row and column offsets to place the total value and a "Total" label directly below the data rows before saving the file.
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
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("D");
            sheet.Cells["B5"].PutValue(40);

            // Define the range for the table (A1:B5)
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 5;         // header + 4 data rows
            int totalColumns = 2;      // two columns (Item, Amount)

            // Add the ListObject (table) and obtain a reference
            int listObjectIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);
            ListObject table = sheet.ListObjects[listObjectIndex];
            table.DisplayName = "SalesTable";

            // Calculate the sum of the Amount column (second column in the table)
            double total = 0;
            for (int r = 0; r < table.DataRange.RowCount; r++)
            {
                // Column offset 1 corresponds to the "Amount" column
                total += sheet.Cells[
                    table.DataRange.FirstRow + r,
                    table.DataRange.FirstColumn + 1].DoubleValue;
            }

            // Insert the total value at the bottom of the table
            int offsetRow = table.DataRange.RowCount; // after last data row
            int amountColumnOffset = 1;               // same column as Amount
            table.PutCellValue(offsetRow, amountColumnOffset, total);
            table.PutCellValue(offsetRow, 0, "Total"); // label in first column

            // Save the workbook
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
