// Title: Insert a DateTime into a ListObject cell using PutCellValue (row offset 2, column offset 3) – C# Aspose.Cells
// Description: Creates a workbook, defines a table (A1:D2), and uses ListObject.PutCellValue with zero‑based offsets to place the date 2023‑12‑31 in the fourth column of the third data row. A built‑in date style (format 14) is applied to the entire column before saving the file as ListObjectDateDemo.xlsx.
// Keywords: Aspose.Cells ListObject PutCellValue | C# insert date into table cell | Aspose.Cells row offset column offset | date formatting Aspose.Cells | Aspose.Cells table example | PutCellValue DateTime | Aspose.Cells C# tutorial
// Common Searches: Aspose.Cells ListObject PutCellValue date example | How to add a DateTime to a table cell with offsets in C# | Apply date format to a column after using PutCellValue | Zero‑based row and column offsets in Aspose.Cells ListObject | C# Aspose.Cells add date to third data row
// Developer Intent: Add a DateTime value to a specific data cell of a ListObject table using row and column offsets and format the column as a date.
// Use Cases: Generating reports that require a timestamp in a specific table column. | Populating a date column in dynamically created worksheets while preserving consistent formatting. | Extending an existing table with date values programmatically for data analysis or export.
// AI Prompts: Provide C# code that creates a workbook, adds a ListObject, and uses PutCellValue to set a DateTime at row offset 2, column offset 3, then formats the column as a date with Aspose.Cells. | Show an Aspose.Cells example that inserts a date into a table cell using zero‑based offsets and applies a built‑in date style to the whole column. | Explain the behavior of ListObject.PutCellValue with row/column offsets and how to apply date formatting to the affected column in C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a workbook, defines a table (A1:D2), and uses ListObject.PutCellValue with zero‑based offsets to place the date 2023‑12‑31 in the fourth column of the third data row. A built‑in date style (format 14) is applied to the entire column before saving the file as ListObjectDateDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with headers (A1:D2)
            sheet.Cells["A1"].PutValue("Col1");
            sheet.Cells["B1"].PutValue("Col2");
            sheet.Cells["C1"].PutValue("Col3");
            sheet.Cells["D1"].PutValue("Col4");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue(2);
            sheet.Cells["C2"].PutValue(3);
            sheet.Cells["D2"].PutValue(4);

            // Add a ListObject (table) that includes the range A1:D2 and has headers
            int tableIndex = sheet.ListObjects.Add(0, 0, 1, 3, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Define the date value to insert
            DateTime dateToInsert = new DateTime(2023, 12, 31);

            // Put the date value at row offset 2 (third data row) and column offset 3 (fourth column)
            // Row offset and column offset are zero‑based within the table (excluding the header row)
            table.PutCellValue(2, 3, dateToInsert);

            // Optional: format the fourth column as a date for proper display
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // Built‑in date format

            // Apply the style to each cell in the fourth data column
            int dataRows = table.DataRange.RowCount;
            for (int r = 0; r < dataRows; r++)
            {
                table.DataRange[r, 3].SetStyle(dateStyle);
            }

            // Save the workbook
            workbook.Save("ListObjectDateDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
