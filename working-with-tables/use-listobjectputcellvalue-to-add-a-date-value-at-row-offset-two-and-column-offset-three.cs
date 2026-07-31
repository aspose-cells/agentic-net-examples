// Title: Aspose.Cells C# – Insert a DateTime into a ListObject cell with PutCellValue (row offset 2, column offset 3)
// Description: This example creates a new workbook, defines a ListObject (table) over A1:D3, and uses ListObject.PutCellValue to place a DateTime value at the specified row and column offsets. It then applies a built‑in date style to the entire Date column and saves the file as an XLSX workbook.
// Keywords: Aspose.Cells ListObject PutCellValue date | C# insert DateTime into table cell | Aspose.Cells format date column | ListObject table cell value example | Aspose.Cells .NET date formatting | Excel table row offset column offset | Aspose.Cells GitHub sample
// Common Searches: How to add a DateTime to a ListObject cell in C# | Aspose.Cells PutCellValue row offset column offset | Formatting date columns after inserting values with Aspose.Cells | C# Aspose.Cells example for inserting dates into tables | GitHub Aspose.Cells ListObject PutCellValue sample
// Developer Intent: Add a DateTime to a specific cell of a ListObject table and format the column as a date.
// Use Cases: Populate transaction dates in dynamically generated sales reports. | Add audit timestamps to rows of a log file exported to Excel. | Update project milestone dates in a schedule table before distribution.
// AI Prompts: Write C# code that uses Aspose.Cells ListObject.PutCellValue to insert a DateTime at row offset 2, column offset 3 and then apply a date style to the column. | Explain how to format a ListObject column as a date after inserting values with PutCellValue in Aspose.Cells for .NET. | Provide a step‑by‑step guide to create a table, insert a date value, style the column, and save the workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // This example creates a new workbook, defines a ListObject (table) over A1:D3, and uses ListObject.PutCellValue to place a DateTime value at the specified row and column offsets. It then applies a built‑in date style to the entire Date column and saves the file as an XLSX workbook.
    public class ListObjectPutCellValueDateDemo
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
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate header row
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["C1"].PutValue("Amount");
            worksheet.Cells["D1"].PutValue("Date"); // Date column

            // Populate data rows
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("Alice");
            worksheet.Cells["C2"].PutValue(100);

            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Bob");
            worksheet.Cells["C3"].PutValue(200);

            // Create a ListObject (table) covering A1:D3 with headers
            int listObjectIndex = worksheet.ListObjects.Add(0, 0, 2, 3, true);
            ListObject listObject = worksheet.ListObjects[listObjectIndex];

            // Insert a date value into the second data row (row offset 1) and fourth column (offset 3)
            DateTime dateToInsert = new DateTime(2023, 5, 15);
            listObject.PutCellValue(1, 3, dateToInsert);

            // Format the Date column to display as a date
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // Built‑in date format (mm-dd-yy)
            // Apply style to the entire Date column within the table
            listObject.ListColumns[3].Range.SetStyle(dateStyle);

            // Save the workbook
            workbook.Save("ListObjectPutCellValueDateDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
