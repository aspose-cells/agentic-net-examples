// Title: Add a DateTime to the fourth column of the third row in an Aspose.Cells ListObject using PutCellValue (C#)
// AI Prompts: Insert a DateTime (e.g., 2023-12-25) into a ListObject at row offset 2 and column offset 3 with the PutCellValue method. | Verify the table has at least four columns and, if necessary, resize the ListObject before writing the date value. | Persist the changes by saving the workbook to an XLSX file after updating the table. | Wrap the insertion logic in a try‑catch block to handle potential runtime errors.
// Common Searches: Aspose.Cells C# how to put a date into a table cell using row and column offsets | Resize ListObject to add extra column before inserting values Aspose.Cells | PutCellValue example with DateTime in an Aspose.Cells ListObject | Insert value at specific offset in Aspose.Cells table C#
// Tags: Aspose.Cells ListObject PutCellValue DateTime | C# Aspose.Cells resize ListObject columns | Aspose.Cells insert value with row offset | Aspose.Cells save workbook to XLSX | Aspose.Cells table column expansion

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The program creates a new workbook, adds a ListObject table, expands it to ensure a fourth column exists, inserts a DateTime (2023‑12‑25) at row offset 2 and column offset 3 using PutCellValue, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define an initial range for the ListObject (e.g., A1:C5)
            int startRow = 0;          // A1 row index
            int startColumn = 0;       // A1 column index
            int rows = 5;
            int cols = 3;              // initially 3 columns (A, B, C)

            // Add a ListObject (table) to the worksheet
            int listObjectIndex = sheet.ListObjects.Add(startRow, startColumn, rows, cols, true);
            ListObject table = sheet.ListObjects[listObjectIndex];
            table.DisplayName = "MyTable";

            // Ensure the table has at least 4 columns (offset 3 = fourth column)
            if (table.ListColumns.Count <= 3)
            {
                // Expand the table to include a fourth column (D)
                // The last parameter indicates whether the table has headers
                table.Resize(startRow, startColumn, rows, 4, true);
            }

            // Date value to insert
            DateTime dateValue = new DateTime(2023, 12, 25);

            // Put the date value at row offset 2 (third row) and column offset 3 (fourth column)
            table.PutCellValue(2, 3, dateValue);

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
