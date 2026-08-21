// Title: C# – Insert a String into an Aspose.Cells ListObject Table Cell using PutCellValue (row & column offsets)
// Description: This example creates a new workbook, defines a header and data range, converts A1:C3 into a ListObject (table), and uses ListObject.PutCellValue with a row offset of 2 and a column offset of 1 to write the string "Charlie" into cell B4. The workbook is then saved as an XLSX file.
// Keywords: Aspose.Cells ListObject PutCellValue | C# insert string into table cell | row offset column offset Aspose.Cells | Aspose.Cells .NET table cell update | Excel table PutCellValue example | Aspose.Cells C# workbook table manipulation
// Common Searches: Aspose.Cells ListObject PutCellValue C# example | how to set a table cell value with rowOffset and columnOffset | insert string into Aspose.Cells table cell | C# Aspose.Cells add value to specific table row | update ListObject cell without expanding range
// Developer Intent: Place a specific string into a designated cell of a ListObject table by providing row and column offsets.
// Use Cases: Add a new employee name to the "Name" column at a calculated row index. | Update a status field in an existing data table based on runtime calculations. | Populate a ListObject with values from an external source by iterating with row offsets.
// AI Prompts: Generate C# code that creates a workbook, adds a ListObject, and uses PutCellValue to insert a string at rowOffset 2, columnOffset 1. | Explain how rowOffset and columnOffset are determined relative to the table header when using ListObject.PutCellValue in Aspose.Cells. | Provide a C# snippet that loops through a collection and inserts each item into successive rows of a ListObject using PutCellValue.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // This example creates a new workbook, defines a header and data range, converts A1:C3 into a ListObject (table), and uses ListObject.PutCellValue with a row offset of 2 and a column offset of 1 to write the string "Charlie" into cell B4. The workbook is then saved as an XLSX file.
    public class ListObjectPutCellValueStringDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate header and data cells that will become a table
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");
            cells["C1"].PutValue("Score");

            cells["A2"].PutValue(1);
            cells["B2"].PutValue("Alice");
            cells["C2"].PutValue(85);

            cells["A3"].PutValue(2);
            cells["B3"].PutValue("Bob");
            cells["C3"].PutValue(92);

            // Add a ListObject (table) covering the range A1:C3, include header row
            int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 2, true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Insert a string value into the table using row and column offsets
            // Row offset = 2 (third data row, zero‑based after the header)
            // Column offset = 1 (second column, i.e., "Name")
            // The value "Charlie" will be placed in cell B4 (row 4, column 2 in the worksheet)
            table.PutCellValue(rowOffset: 2, columnOffset: 1, value: "Charlie");

            // Save the workbook to an XLSX file
            string outputPath = "ListObjectPutCellValueStringDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
