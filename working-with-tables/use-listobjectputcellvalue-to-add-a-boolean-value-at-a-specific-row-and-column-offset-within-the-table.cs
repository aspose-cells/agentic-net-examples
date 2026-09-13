// Title: How to use ListObject.PutCellValue to insert a Boolean into a specific row and column of an Excel table with Aspose.Cells in C#
// AI Prompts: Generate C# code that creates a workbook, adds a ListObject table, and calls ListObject.PutCellValue to set a true Boolean at row offset 1 and column offset 1. | Show a step‑by‑step example of updating the "Flag" column in the second data row of an Aspose.Cells ListObject using PutCellValue. | Provide a complete C# snippet that saves the workbook after inserting a Boolean value into a table cell via ListObject.PutCellValue.
// Common Searches: Aspose.Cells C# put boolean value in ListObject table cell by row and column offset | How to set true in a specific cell of an Excel table using Aspose.Cells ListObject.PutCellValue | C# example for inserting a Boolean into the Flag column of a ListObject | PutCellValue method row offset column offset Aspose.Cells tutorial | Add boolean data to an Excel table created with Aspose.Cells in C#
// Tags: Aspose.Cells ListObject.PutCellValue boolean | C# insert boolean into Excel table | ListObject row offset column offset | Aspose.Cells create and save workbook with table | update Flag column Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// Demonstrates creating a workbook, defining a ListObject table, and using ListObject.PutCellValue to place a true Boolean in the second data row of the Flag column, then saving the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data (header + rows). Column B (Flag) is initially empty.
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Flag");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["A4"].PutValue(3);

            // Define the range that will become the ListObject (table)
            int firstRow = 0;      // Row index for A1
            int firstColumn = 0;   // Column index for A
            int totalRows = 4;     // A1:B4 (including header)
            int totalColumns = 2;  // Two columns: ID and Flag
            bool hasHeaders = true;

            // Add the ListObject to the worksheet
            int listObjectIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, hasHeaders);
            ListObject listObject = sheet.ListObjects[listObjectIndex];
            listObject.DisplayName = "SampleTable";

            // Put a boolean value (true) at a specific row and column offset within the table
            // Row offset 1 => second data row (originally row 3 in the sheet)
            // Column offset 1 => second column of the table (Flag column)
            listObject.PutCellValue(1, 1, true);

            // Save the workbook to a file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
