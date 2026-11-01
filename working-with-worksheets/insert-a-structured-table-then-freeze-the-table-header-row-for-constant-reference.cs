// Title: Add a ListObject structured table, apply a built‑in style, and freeze the header row in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a ListObject from a range, sets TableStyleMedium9, and freezes the first worksheet row. | Show how to use the Aspose.Cells FreezePanes overload to keep a table header visible while scrolling in a generated .xlsx file.
// Common Searches: asp.net core generate Excel file with structured table and frozen header using Aspose.Cells | c# Aspose.Cells create ListObject from range and apply built‑in table style | how to freeze first row of a worksheet that contains a ListObject in Aspose.Cells | Aspose.Cells FreezePanes example with table headers in C#
// Tags: Aspose.Cells structured table creation C# | built‑in table style application Aspose.Cells | Excel header row freezing Aspose.Cells | convert cell range to table Aspose.Cells | C# generate .xlsx with styled table and frozen panes

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The sample creates a new workbook, fills cells A1:C4 with headers and data, converts that range into a ListObject (structured table) with headers, applies the built‑in TableStyleMedium9, freezes the header row using the FreezePanes overload, and saves the workbook as StructuredTable.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Populate data including header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);

            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");
            sheet.Cells["C4"].PutValue(78);

            // Define the range that will become a structured table
            var tableRange = sheet.Cells.CreateRange("A1:C4");

            // Add a ListObject (structured table) to the defined range, specifying that the range has headers
            int listObjectIndex = sheet.ListObjects.Add(
                tableRange.FirstRow,
                tableRange.FirstColumn,
                tableRange.RowCount,
                tableRange.ColumnCount,
                true); // hasHeaders

            var listObject = sheet.ListObjects[listObjectIndex];
            // Apply a built‑in table style
            listObject.TableStyleType = TableStyleType.TableStyleMedium9;

            // Freeze the header row so it stays visible while scrolling
            // Use the overload with four parameters (row, column, totalRows, totalColumns)
            sheet.FreezePanes(1, 0, sheet.Cells.MaxDataRow + 1, sheet.Cells.MaxDataColumn + 1);

            // Save the workbook
            string outputPath = "StructuredTable.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
