// Title: C# – Create an Excel Table from a Cell Range and Assign a Custom Name with Aspose.Cells
// Description: Learn how to use Aspose.Cells for .NET to convert a range (A1:B3) into a ListObject table, set its DisplayName (e.g., "EmployeeTable"), apply a built‑in style, and save the workbook as an .xlsx file.
// Keywords: Aspose.Cells create table from range | C# ListObject custom name | Aspose.Cells set table DisplayName | apply table style Aspose.Cells | save workbook as xlsx .NET | convert cell range to Excel table | Aspose.Cells TableStyleMedium9
// Common Searches: convert cell range to table Aspose.Cells .NET | set custom table name in Aspose.Cells | apply built‑in table style with Aspose.Cells | add ListObject with headers using Aspose.Cells | save Aspose.Cells workbook as xlsx
// Developer Intent: Create a ListObject table from a specified range, give it a custom name, optionally style it, and export the workbook.
// Use Cases: Generate a named employee list table for reporting. | Programmatically build styled tables from dynamic data ranges. | Automate Excel workbook creation with named tables to simplify downstream analysis.
// AI Prompts: Write C# code with Aspose.Cells that creates a table from range A1:C10, names it "SalesData", applies TableStyleMedium9, and saves as an .xlsx file. | Show how to add a ListObject, set its DisplayName, apply a built‑in style, and persist the workbook using Aspose.Cells for .NET. | Explain how to retrieve the index of a newly added ListObject, rename the table, and confirm the applied style.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

// Learn how to use Aspose.Cells for .NET to convert a range (A1:B3) into a ListObject table, set its DisplayName (e.g., "EmployeeTable"), apply a built‑in style, and save the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (including header row)
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");
            cells["A2"].PutValue(1);
            cells["B2"].PutValue("John");
            cells["A3"].PutValue(2);
            cells["B3"].PutValue("Mary");

            // Define the range that will be converted to a table
            AsposeRange sourceRange = cells.CreateRange("A1", "B3");

            // Add a ListObject (table) using the range coordinates
            int tableIndex = sheet.ListObjects.Add(
                sourceRange.FirstRow,
                sourceRange.FirstColumn,
                sourceRange.FirstRow + sourceRange.RowCount - 1,
                sourceRange.FirstColumn + sourceRange.ColumnCount - 1,
                true); // true indicates the range has headers

            ListObject table = sheet.ListObjects[tableIndex];

            // Assign a custom name to the table
            table.DisplayName = "EmployeeTable";

            // Optional: apply a built‑in table style
            table.TableStyleName = "TableStyleMedium9";

            // Save the workbook
            string outputPath = "CreatedTable.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
