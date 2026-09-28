// Title: Add a totals row to a ListObject table and count text entries using Aspose.Cells for .NET (C#)
// AI Prompts: Create a ListObject on a worksheet, enable its totals row, and write the number of rows in a text column to the totals row with Aspose.Cells. | Insert a SUM formula for a numeric column in the totals row of an Aspose.Cells table while displaying the count of a string column. | Generate an Excel file in C# that contains sample data, a table, and a totals row showing a row count for the Name column and a sum for the Amount column using Aspose.Cells.
// Common Searches: Aspose.Cells C# add totals row to ListObject and count string column | how to set count aggregation for a text field in an Excel table using Aspose.Cells | C# Aspose.Cells totals row sum formula for numeric column and row count for text column
// Tags: listobject totals row Aspose.Cells | text column count aggregation Aspose.Cells | numeric column sum formula totals row C# | create Excel table with totals row Aspose.Cells | Aspose.Cells set totals row values programmatically

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example creates a workbook, adds sample data, defines a ListObject covering A1:B4, enables the totals row, writes the row count of the text 'Name' column into the totals row, inserts a SUM formula for the numeric 'Amount' column, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (Name column is text)
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["A4"].PutValue("Charlie");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a ListObject (table) that includes the data range A1:B4
            int listIndex = sheet.ListObjects.Add("A1", "B4", true);
            ListObject listObject = sheet.ListObjects[listIndex];

            // Enable the totals row
            listObject.ShowTotals = true;

            // Determine the totals row index (first row after the data range)
            int totalsRowIndex = listObject.DataRange.FirstRow + listObject.DataRange.RowCount;

            // Set count for the "Name" column (simple count of rows)
            int nameColumnIndex = listObject.DataRange.FirstColumn;
            int rowCount = listObject.DataRange.RowCount;
            sheet.Cells[totalsRowIndex, nameColumnIndex].PutValue(rowCount);

            // Set sum for the "Amount" column using a formula
            int amountColumnIndex = nameColumnIndex + 1;
            string amountColumnLetter = CellsHelper.ColumnIndexToName(amountColumnIndex);
            int dataStartRow = listObject.DataRange.FirstRow + 2; // data starts at row 2 (A2)
            int dataEndRow = dataStartRow + rowCount - 1;
            string sumFormula = $"=SUM({amountColumnLetter}{dataStartRow}:{amountColumnLetter}{dataEndRow})";
            sheet.Cells[totalsRowIndex, amountColumnIndex].Formula = sumFormula;

            // Determine output file path
            string outputPath = "output.xlsx";

            // Save the workbook (overwrite if it already exists)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
