// Title: Insert a Boolean value into a specific row and column of an Aspose.Cells ListObject table using C# PutCellValue
// Common Searches: how to write true to a ListObject cell by row and column offset in Aspose.Cells C# | Aspose.Cells C# set boolean value in Excel table without using a cell address | using ListObject.PutCellValue to update a flag column in a generated worksheet | add boolean data to a table column in Aspose.Cells workbook programmatically | C# Aspose.Cells example for inserting true into the third data row of a table
// Tags: Aspose.Cells ListObject PutCellValue example | C# insert boolean into Excel table | zero‑based row offset table cell update | Aspose.Cells write true to table column | programmatic Excel table modification C#

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The sample creates a new workbook, defines a two‑column ListObject with a header row, and uses ListObject.PutCellValue with zero‑based row and column offsets to write the Boolean value true into the 'Flag' column of the third data row, then saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data that will become the table
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Flag");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["A4"].PutValue(3);

            // Define the range for the ListObject (table) – includes header row
            int firstRow = 0;          // zero‑based index for row 1 (A1)
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 4;         // rows A1:B4
            int totalColumns = 2;      // columns A and B

            // Add the ListObject to the worksheet
            int listObjectIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject table = sheet.ListObjects[listObjectIndex];
            table.DisplayName = "MyTable";

            // Add a boolean value using PutCellValue
            // rowOffset: zero‑based offset from the first data row (excluding header)
            // columnOffset: zero‑based offset within the table columns
            int rowOffset = 2;      // third data row (row 4 in the sheet)
            int columnOffset = 1;   // second column of the table ("Flag")
            table.PutCellValue(rowOffset, columnOffset, true);

            // Determine output path and ensure the directory exists
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
