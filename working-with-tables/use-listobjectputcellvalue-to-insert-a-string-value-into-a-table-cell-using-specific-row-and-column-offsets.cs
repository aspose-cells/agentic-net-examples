// Title: Insert a text value into a specific table cell using ListObject.PutCellValue with row and column offsets in Aspose.Cells for .NET
// AI Prompts: Show how to create a ListObject covering a range, then use PutCellValue(rowOffset, columnOffset, value) to write a string into the second data row of the table in C#. | Provide a C# snippet that adds an Excel table, calculates row and column offsets, and updates a cell with a custom text using Aspose.Cells ListObject.PutCellValue. | Generate code that demonstrates saving the workbook after inserting a value into a table cell by offset with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# putcellvalue into table cell using row and column offsets | how to update a specific data row in an Excel table with Aspose.Cells ListObject | C# example of inserting text into a ListObject cell by offset | using ListObject.PutCellValue to modify table data in Aspose.Cells for .NET | Aspose.Cells putcellvalue method offset parameters explanation
// Tags: ListObject.PutCellValue with row offset | update Excel table cell Aspose.Cells C# | insert string into table cell using Aspose.Cells | Aspose.Cells table cell modification by offset | C# Aspose.Cells workbook save after table edit

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The program creates a new workbook, defines a ListObject (table) over range A1:B3, then uses ListObject.PutCellValue with specified row and column offsets to insert the string "Inserted Value" into the second data row of the table, and finally saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some initial data to define the table range
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue("Row1Col1");
            sheet.Cells["B2"].PutValue("Row1Col2");
            sheet.Cells["A3"].PutValue("Row2Col1");
            sheet.Cells["B3"].PutValue("Row2Col2");

            // Create a ListObject (table) covering the range A1:B3
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 3;         // rows including header
            int totalColumns = 2;      // columns
            int tableIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);

            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "MyTable";

            // Define row and column offsets relative to the first data row of the table
            int rowOffset = 1;    // second data row (zero‑based)
            int columnOffset = 0; // first column of the table

            // Insert a string value into the specified cell using PutCellValue
            string newValue = "Inserted Value";
            table.PutCellValue(rowOffset, columnOffset, newValue);

            // Determine output file path
            string outputPath = "output.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
