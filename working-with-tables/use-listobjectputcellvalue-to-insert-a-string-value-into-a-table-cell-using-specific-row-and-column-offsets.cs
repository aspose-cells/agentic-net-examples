// Title: C# – Set a table cell with ListObject.PutCellValue using row and column offsets in Aspose.Cells
// Description: Demonstrates how to create a workbook, define a table (ListObject) over range A1:C3, and replace the value in the second column of the first data row (cell B2) with "Charlie" by calling ListObject.PutCellValue(rowOffset, columnOffset, value). The file is saved as an XLSX document.
// Keywords: Aspose.Cells ListObject PutCellValue C# | update Excel table cell by offset | Aspose.Cells table cell modification | C# Excel table row column offset | Aspose.Cells PutCellValue example
// Common Searches: Aspose.Cells ListObject PutCellValue example | how to change a table cell using row and column offsets in .NET | update specific column in Excel table Aspose.Cells | C# set value in ListObject cell by offset | Aspose.Cells replace cell value in table
// Developer Intent: Update a single cell inside an Excel table by specifying its row and column offsets rather than an absolute address.
// Use Cases: Correct a mistaken entry in a table column without hard‑coding the cell reference. | Iterate through newly added rows and fill each column using offset indices. | Perform bulk edits on a table by looping over row/column offsets and calling PutCellValue.
// AI Prompts: Write C# code that updates multiple ListObject cells by looping over row and column offsets with PutCellValue. | Show how to validate rowOffset and columnOffset before calling ListObject.PutCellValue and handle out‑of‑range errors. | Explain how to read offset values from a configuration file and apply them to modify an Aspose.Cells table.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsListObjectPutCellValueDemo
{
    // Demonstrates how to create a workbook, define a table (ListObject) over range A1:C3, and replace the value in the second column of the first data row (cell B2) with "Charlie" by calling ListObject.PutCellValue(rowOffset, columnOffset, value). The file is saved as an XLSX document.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate header row for the table
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");
            cells["C1"].PutValue("Score");

            // Add some initial data rows
            cells["A2"].PutValue(1);
            cells["B2"].PutValue("Alice");
            cells["C2"].PutValue(85);

            cells["A3"].PutValue(2);
            cells["B3"].PutValue("Bob");
            cells["C3"].PutValue(92);

            // Create a ListObject (table) that covers the range A1:C3, including the header
            int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 2, true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Insert a string value into the table using row and column offsets
            // Row offset = 1 refers to the first data row (row index 1 in the table, i.e., Excel row 2)
            // Column offset = 1 refers to the second column of the table (the "Name" column)
            table.PutCellValue(1, 1, "Charlie"); // Updates cell B2 with "Charlie"

            // Save the workbook to a file
            workbook.Save("ListObjectPutCellValueDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
