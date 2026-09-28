// Title: Insert a multi‑line string into a ListObject table cell and enable text wrapping using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells ListObject.PutCellValue to write a newline‑separated string into a specific table cell and apply a wrapped‑text style in C#. | Create a ListObject table, add multi‑line text to a cell, set IsTextWrapped = true on the corresponding worksheet cell, then save the workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# how to add newline characters to a ListObject table cell | Insert multi line text into an Aspose.Cells table cell example | Enable text wrapping for a cell after inserting multi‑line string in Aspose.Cells | Saving an XLSX workbook after writing multi‑line data to a ListObject in C# | Aspose.Cells ListObject insert multi‑line string with text wrap
// Tags: Aspose.Cells ListObject multi-line cell | Aspose.Cells text wrap style C# | C# insert newline string into table cell | Aspose.Cells create and style table | Aspose.Cells save workbook xlsx

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample creates a new workbook, adds a ListObject table, defines a multi‑line string using '\n', inserts it into the second row‑second column of the table with PutCellValue, applies a style with IsTextWrapped set to true to the corresponding worksheet cell, and saves the file as MultiLineTable.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Define the range for the ListObject (table) – starts at A1, 3 rows x 2 columns
            int firstRow = 0;          // A1 row index (zero‑based)
            int firstColumn = 0;       // A1 column index (zero‑based)
            int totalRows = 3;
            int totalColumns = 2;

            // Add the ListObject (table) to the worksheet
            int listObjectIndex = sheet.ListObjects.Add(firstRow, firstColumn,
                                                        firstRow + totalRows,
                                                        firstColumn + totalColumns, true);
            ListObject table = sheet.ListObjects[listObjectIndex];
            table.DisplayName = "MyTable";

            // Multi‑line string to insert (use \n for line breaks)
            string multiLineText = "Line 1\nLine 2\nLine 3";

            // Put the multi‑line string into the second row, second column of the table
            // (row and column indices are zero‑based relative to the table)
            table.PutCellValue(1, 1, multiLineText);

            // Enable text wrapping for the target cell so the newlines are visible
            Style wrapStyle = workbook.CreateStyle();
            wrapStyle.IsTextWrapped = true;

            // Calculate the absolute cell address within the worksheet
            int targetRow = table.StartRow + 1;      // table.StartRow is the first row of the table
            int targetColumn = table.StartColumn + 1;
            sheet.Cells[targetRow, targetColumn].SetStyle(wrapStyle);

            // Define output file path
            string outputPath = "MultiLineTable.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
