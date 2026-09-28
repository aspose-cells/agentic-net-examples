// Title: How to create a dynamic Excel table with an auto‑expanding named range using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to add a ListObject table, insert new rows, and automatically resize the table so the named range expands. | Show the steps to define a named range for an Excel table in Aspose.Cells and update it after appending data rows. | Provide a snippet that creates a workbook, builds a table with headers, adds a row, and calls Resize to keep the range dynamic.
// Common Searches: Aspose.Cells C# create ListObject with dynamic named range that grows when rows are added | Resize Excel table programmatically after inserting rows using Aspose.Cells .NET | How to make an Excel table auto‑expand its range in Aspose.Cells C# example
// Tags: Aspose.Cells ListObject dynamic range | C# resize Excel table Aspose | auto‑expand Excel table named range .NET | create Excel table with headers Aspose.Cells | add data row to Aspose.Cells ListObject

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a new workbook, defines a ListObject table with headers, adds an extra data row, resizes the table to include the new row, and saves the file as DynamicTable.xlsx, demonstrating a dynamic named range that expands automatically.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Add header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            // Add initial data rows
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(90);

            // Define the initial range for the table (header + 2 data rows)
            int firstRow = 0;          // zero‑based index for row 1 (header)
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 3;         // 1 header + 2 data rows
            int totalColumns = 3;      // ID, Name, Score

            // Create a ListObject (Excel table) which provides a dynamic named range
            // The last parameter 'hasHeaders' indicates that the first row contains headers
            int tableIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);

            ListObject table = sheet.ListObjects[tableIndex];
            table.ShowHeaderRow = true;
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Add a new row below the current table; the table will expand automatically
            int newRowIdx = sheet.Cells.MaxDataRow + 1; // first empty row after existing data
            sheet.Cells[newRowIdx, 0].PutValue(3);      // ID
            sheet.Cells[newRowIdx, 1].PutValue("Charlie"); // Name
            sheet.Cells[newRowIdx, 2].PutValue(78);    // Score

            // Resize the table to include the newly added row (hasHeaders = true)
            int newTotalRows = sheet.Cells.MaxDataRow + 1; // total rows including header
            table.Resize(firstRow, firstColumn, newTotalRows, totalColumns, true);

            // Ensure output directory exists
            string outputPath = "DynamicTable.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
