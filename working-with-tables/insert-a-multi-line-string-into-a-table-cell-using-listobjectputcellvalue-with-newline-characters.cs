// Title: How to insert a multi‑line string into a ListObject table cell and enable text wrapping with Aspose.Cells for C#
// AI Prompts: Insert a newline‑separated string into a specific ListObject cell using PutCellValue and apply a wrapped style in C#. | Write C# code that adds multi‑line text to a table cell in Aspose.Cells and turns on text wrapping via StyleFlag. | Demonstrate setting text wrap for a ListObject data cell after putting a multi‑line value with the Aspose.Cells API.
// Common Searches: Aspose.Cells C# put multi line text into ListObject cell with newline characters | Enable text wrap for a table cell after using PutCellValue in Aspose.Cells | How to use StyleFlag to wrap text in an Excel table cell with Aspose.Cells C# | Insert multi‑line string into Excel table using Aspose.Cells ListObject PutCellValue | C# Aspose.Cells example for multiline cell content in a ListObject
// Tags: Aspose.Cells ListObject PutCellValue multiline | Aspose.Cells text wrap StyleFlag C# | Excel table cell multiline Aspose.Cells | C# insert newline string into ListObject cell | Aspose.Cells table cell style wrapping

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;

// The example creates a workbook, adds a ListObject table, inserts a newline‑separated string into the first data cell using PutCellValue, applies a text‑wrapped style with StyleFlag, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for the table (A1:B2) – includes header row
            int firstRow = 0;      // A1 row index
            int firstColumn = 0;   // A1 column index
            int totalRows = 2;     // header + 1 data row
            int totalColumns = 2;  // two columns

            // Add a ListObject (table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn,
                                                   firstRow + totalRows,
                                                   firstColumn + totalColumns, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "MyTable";

            // Multi‑line string to insert (use \n for line breaks)
            string multiLine = "Line 1\nLine 2\nLine 3";

            // Put the value into the first data row, first column of the table
            // Row and column are zero‑based relative to the data range (excluding header)
            int dataRow = 0;   // first data row inside the table
            int dataColumn = 0;
            table.PutCellValue(dataRow, dataColumn, multiLine);

            // Enable text wrapping for the target cell so newlines are visible
            Cell targetCell = sheet.Cells[table.DataRange.FirstRow + dataRow,
                                         table.DataRange.FirstColumn + dataColumn];
            Style wrapStyle = workbook.CreateStyle();
            wrapStyle.IsTextWrapped = true;

            // Apply the style to the cell (apply all style attributes)
            StyleFlag flag = new StyleFlag { All = true };
            targetCell.SetStyle(wrapStyle, flag);

            // Save the workbook
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
