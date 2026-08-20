// Title: C# – Create a Dynamic ListObject (Excel Table) and Auto‑Expand It with Aspose.Cells
// Description: Demonstrates how to build a workbook, define an initial range (A1:B3), add a ListObject (table) with a name and style, insert new rows, calculate the new end row, and call ListObject.Resize to automatically expand the table before saving the file.
// Keywords: Aspose.Cells | C# ListObject | dynamic Excel table | ListObject.Resize | auto expand table | add rows to Aspose.Cells table | create table from range | Excel table automation .NET | Aspose.Cells table resizing | programmatic Excel table expansion
// Common Searches: Aspose.Cells create table from range C# | Resize ListObject after inserting rows Aspose.Cells | How to auto expand Excel table using Aspose.Cells | C# Aspose.Cells dynamic ListObject example | Add rows to Aspose.Cells table programmatically
// Developer Intent: Programmatically create an Excel table from a range and have it grow automatically as new rows are added.
// Use Cases: Generate an employee roster table that expands whenever new staff records are appended. | Import a CSV into a worksheet, convert it to a ListObject, then add monthly data rows without recreating the table. | Build a financial report where each quarter’s rows are inserted and the table range updates automatically.
// AI Prompts: Show C# code that creates a ListObject from a range and automatically resizes it after inserting rows using Aspose.Cells. | Explain how to calculate the new end row for a ListObject when multiple rows are added in Aspose.Cells. | Provide a step‑by‑step guide to auto‑expand an Excel table with Aspose.Cells .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Demonstrates how to build a workbook, define an initial range (A1:B3), add a ListObject (table) with a name and style, insert new rows, calculate the new end row, and call ListObject.Resize to automatically expand the table before saving the file.
    public class DynamicListObjectDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some initial data (including headers)
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Alice");
                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Bob");

                // Define a dynamic range that currently covers the data above
                // The range will be used to create the ListObject (table)
                string startCell = "A1";
                string endCell = "B3"; // initial end cell

                // Add the ListObject (table) to the worksheet using the string‑based Add method
                int tableIndex = sheet.ListObjects.Add(startCell, endCell, true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Optional: give the table a friendly name and a style
                table.DisplayName = "PeopleTable";
                table.TableStyleType = TableStyleType.TableStyleMedium2;

                // -----------------------------------------------------------------
                // Simulate adding new rows to the worksheet below the current table
                // -----------------------------------------------------------------
                // Insert three new rows after the existing data
                sheet.Cells.InsertRows(4, 3, true); // insert at row index 4 (zero‑based)

                // Populate the newly inserted rows with data
                sheet.Cells["A4"].PutValue(3);
                sheet.Cells["B4"].PutValue("Charlie");
                sheet.Cells["A5"].PutValue(4);
                sheet.Cells["B5"].PutValue("Diana");
                sheet.Cells["A6"].PutValue(5);
                sheet.Cells["B6"].PutValue("Eve");

                // -----------------------------------------------------------------
                // Expand the ListObject to include the newly added rows.
                // This mimics the automatic expansion behavior of Excel tables.
                // -----------------------------------------------------------------
                // Calculate the new end row based on the current data range
                int newEndRow = table.EndRow + 3; // we added three rows

                // Resize the table to the new range (same start column/row, new end row)
                table.Resize(table.StartRow, table.StartColumn, newEndRow, table.EndColumn, true);

                // Save the workbook
                string outputPath = "DynamicListObjectDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            DynamicListObjectDemo.Run();
        }
    }
}
