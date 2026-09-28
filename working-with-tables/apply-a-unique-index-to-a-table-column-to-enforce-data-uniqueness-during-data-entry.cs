// Title: Implement manual unique‑value enforcement for an Excel table column using Aspose.Cells in C# (workaround for missing unique index support)
// AI Prompts: Generate C# code with Aspose.Cells that creates a worksheet, adds a ListObject named Employees, and throws an exception when a duplicate Email value is inserted. | Show how to extend the Aspose.Cells example to check for existing ID values before appending a new row, returning an error if the ID already exists. | Provide a reusable C# method that uses Aspose.Cells to validate that a specified column in a ListObject contains only unique values and integrates it into the data‑entry workflow.
// Common Searches: aspocells enforce unique values in a ListObject column c# | c# prevent duplicate rows in Excel table using Aspose.Cells | how to simulate unique index on Excel column with Aspose.Cells | manual uniqueness check before adding data to Aspose.Cells table | aspocells unique constraint workaround for email column
// Tags: Aspose.Cells custom unique validation | C# Excel ListObject duplicate detection | Aspose.Cells manual unique index workaround | Excel table unique column enforcement Aspose | Aspose.Cells data entry uniqueness check

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example creates a new workbook, adds a worksheet named 'Data', defines an Employees ListObject, applies a table style, and notes that Aspose.Cells lacks built‑in unique index support. It suggests implementing custom validation to reject duplicate values before inserting rows and saves the file as UniqueIndexExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Add header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Email");

            // Add sample data rows
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue("alice@example.com");

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue("bob@example.com");

            // Define the range of the table (including header)
            int firstRow = 0;      // zero‑based index for row 1
            int firstColumn = 0;   // zero‑based index for column A
            int lastRow = 2;       // row 3 (header + 2 data rows)
            int lastColumn = 2;    // column C

            // Add a ListObject (Excel table) to the defined range; hasHeaders = true
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, lastRow, lastColumn, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Set the display name of the table (used as the table name in Excel)
            table.DisplayName = "Employees";

            // Apply a built‑in table style
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // NOTE: Adding a unique index on a column is not supported in the current Aspose.Cells version.
            // If needed, implement validation logic manually when inserting data.

            // Save the workbook to a file
            string outputPath = "UniqueIndexExample.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
