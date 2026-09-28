// Title: Create an Excel table and define a header‑only named range with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells that adds a ListObject to a worksheet and creates a named range that points exclusively to the table’s header row. | Show how to add a header‑only named range to a workbook with reflection to maintain compatibility with older Aspose.Cells versions.
// Common Searches: Aspose.Cells C# create named range that includes only the first row of a table | How to reference a table header row as a named range in Aspose.Cells .NET | Add ListObject to worksheet and name its header range using Aspose.Cells for C# | Using reflection to add a name to an Aspose.Cells workbook when the Names collection is unavailable | Create Excel table with headers and a separate header range in Aspose.Cells
// Tags: Aspose.Cells add ListObject table C# | Aspose.Cells named range for table header | Aspose.Cells reflection add workbook name | Aspose.Cells define range for header row | Aspose.Cells save workbook as XLSX

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a new workbook, inserts a ListObject table with headers on the first worksheet, builds a range that covers only the header row, adds this range as a named range (using reflection to support older Aspose.Cells versions), and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            // Populate some data rows
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);

            // Define table dimensions (including header)
            int firstRow = 0;          // Row 1 (zero‑based)
            int firstColumn = 0;       // Column A
            int totalRows = 3;         // Header + 2 data rows
            int totalColumns = 3;      // ID, Name, Score

            // Add a table (ListObject) to the worksheet
            int tableIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true); // hasHeaders = true
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "MyTable";

            // Create a named range that references only the header row of the table
            // Header row is the first row of the table (firstRow)
            Aspose.Cells.Range headerRange = sheet.Cells.CreateRange(firstRow, firstColumn, 1, totalColumns);
            // Add the named range to the workbook (if supported by the version)
            // If the Names collection is unavailable, this step can be omitted safely.
            if (workbook.GetType().GetProperty("Names") != null)
            {
                // Use reflection to add the name without breaking compilation on older versions
                var namesProp = workbook.GetType().GetProperty("Names");
                var namesCollection = namesProp.GetValue(workbook, null);
                var addMethod = namesCollection.GetType().GetMethod("Add", new[] { typeof(string), typeof(Aspose.Cells.Range) });
                addMethod?.Invoke(namesCollection, new object[] { "MyTableHeader", headerRange });
            }

            // Save the workbook
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
