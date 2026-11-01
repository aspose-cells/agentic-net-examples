// Title: How to use ListObject.PutCellValue with row and column offsets to insert a string into an Aspose.Cells table in C#
// AI Prompts: Insert the string "Passed" into the third data row and third column of a ListObject by calling ListObject.PutCellValue with zero‑based offsets in C#. | Create a ListObject for range A1:C5 and modify a cell using ListObject.PutCellValue by specifying row and column offsets.
// Common Searches: Aspose.Cells C# putcellvalue offset example for table | how to update a specific cell in an Excel ListObject using Aspose.Cells | apply offset indices with ListObject.PutCellValue in C# | write a value to the cell at row 3 column 3 of an Aspose.Cells table | Aspose.Cells ListObject PutCellValue row offset 2 column offset 2
// Tags: ListObject.PutCellValue with row offset | Aspose.Cells table cell update C# | Excel table offset insertion Aspose.Cells | C# Aspose.Cells ListObject value assignment | PutCellValue offset usage in Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The program creates a new workbook, defines a ListObject covering A1:C5, and uses ListObject.PutCellValue with row offset 2 and column offset 2 to write the string "Passed" into the third data row and third column of the table, then saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data to form a table (A1:C5)
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            for (int i = 2; i <= 5; i++)
            {
                sheet.Cells[i - 1, 0].PutValue(i - 1);               // ID
                sheet.Cells[i - 1, 1].PutValue($"Item{i - 1}");    // Name
                sheet.Cells[i - 1, 2].PutValue(0);                 // Score
            }

            // Add a ListObject (table) covering the range A1:C5
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 5;
            int totalColumns = 3;

            int tableIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);

            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "MyTable";

            // Insert a string value into a specific cell of the table using offsets
            // Row offset = 2 (third data row, zero‑based, after header)
            // Column offset = 2 (third column of the table)
            int rowOffset = 2;
            int columnOffset = 2;
            table.PutCellValue(rowOffset, columnOffset, "Passed");

            // Save the workbook (save rule)
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
