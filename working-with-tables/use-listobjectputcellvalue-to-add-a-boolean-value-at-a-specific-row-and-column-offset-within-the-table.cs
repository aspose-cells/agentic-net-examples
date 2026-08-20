// Title: Set a Boolean in an Aspose.Cells ListObject Table with ListObject.PutCellValue (C#)
// Description: Demonstrates how to create a workbook, define a table with headers, add sample rows, and use ListObject.PutCellValue(rowOffset, columnOffset, true) to write a boolean value into a specific cell of the table before saving the file.
// Keywords: Aspose.Cells ListObject PutCellValue | C# boolean table cell | Aspose.Cells update cell by offset | Set true false in Aspose table | Aspose.Cells .NET ListObject example
// Common Searches: Aspose.Cells PutCellValue boolean example | How to write true/false to a ListObject cell in C# | Update specific table cell using rowOffset columnOffset Aspose | Set boolean value in Aspose.Cells table programmatically
// Developer Intent: Insert a boolean value into a designated row and column of an Aspose.Cells ListObject using the PutCellValue method.
// Use Cases: Mark records as active/inactive in a generated report. | Apply business‑logic flags to rows after data import. | Initialize a status column when building worksheets dynamically.
// AI Prompts: Show how to use ListObject.PutCellValue to write a DateTime value at a given offset. | Provide a loop that updates multiple boolean cells in a ListObject with PutCellValue. | Explain how to convert a ListObject row/column offset to an A1 cell address before calling PutCellValue.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, define a table with headers, add sample rows, and use ListObject.PutCellValue(rowOffset, columnOffset, true) to write a boolean value into a specific cell of the table before saving the file.
    public class ListObjectPutBooleanDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add header row for the table
                worksheet.Cells["A1"].PutValue("ID");
                worksheet.Cells["B1"].PutValue("IsActive");

                // Add some initial data rows
                worksheet.Cells["A2"].PutValue(1);
                worksheet.Cells["B2"].PutValue(false);
                worksheet.Cells["A3"].PutValue(2);
                worksheet.Cells["B3"].PutValue(true);

                // Create a ListObject (table) that covers the data range A1:B3
                // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
                int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Update the cell at row offset 1 (second data row) and column offset 1 (second column)
                // Set the boolean value to true using PutCellValue
                table.PutCellValue(rowOffset: 1, columnOffset: 1, value: true);

                // Save the workbook to a file
                string outputPath = "ListObjectPutCellValueBooleanDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
