// Title: Insert a DateTime into a specific cell of an Aspose.Cells ListObject using PutCellValue (row offset 2, column offset 3) in C#
// Common Searches: how to set a date value in an Aspose.Cells table cell with PutCellValue C# | Aspose.Cells ListObject PutCellValue row offset column offset example | C# code to add a DateTime to a table cell at row 3 column 4 using Aspose.Cells | using Aspose.Cells to write a date into an Excel table at specific offsets | Aspose.Cells PutCellValue date insertion into ListObject in .NET
// Tags: Aspose.Cells ListObject PutCellValue date | C# set DateTime in Excel table using Aspose.Cells | row offset column offset cell update Aspose.Cells | Excel table manipulation with Aspose.Cells API | Aspose.Cells .NET write date to ListObject cell

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

// The example creates a new workbook, adds a ListObject covering A1:D5, and uses ListObject.PutCellValue with a row offset of 2 and column offset of 3 to write the DateTime 2023‑12‑25 into that table cell, then saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for the table (A1:D5)
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 5;         // rows 1‑5
            int totalColumns = 4;      // columns A‑D

            // Add a ListObject (table) to the worksheet
            int listObjectIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows,
                firstColumn + totalColumns,
                true);
            ListObject listObject = sheet.ListObjects[listObjectIndex];
            listObject.DisplayName = "MyTable";

            // Row offset two and column offset three (zero‑based inside the table)
            int rowOffset = 2;      // third row inside the table
            int columnOffset = 3;   // fourth column inside the table

            // Put a date value at the specified cell within the table
            DateTime dateValue = new DateTime(2023, 12, 25);
            listObject.PutCellValue(rowOffset, columnOffset, dateValue);

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
