// Title: Set Excel table header cells with ListObject.PutCellValue using row and column offsets in Aspose.Cells for .NET
// Common Searches: how to use ListObject.PutCellValue to write header values in an Aspose.Cells table C# | populate Excel table column names with row and column offsets in Aspose.Cells .NET | Aspose.Cells C# example for adding a ListObject and setting header cells | using PutCellValue to assign table headers in a workbook created with Aspose.Cells | C# code to create a table range and set header text via ListObject offsets
// Tags: listobject putcellvalue header cells | aspocells create table range c# | excel table column names aspocells | c# set table header offsets | aspocells workbook save output.xlsx

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables; // Required for ListObject

// The sample creates a new Workbook, adds a ListObject covering A1:C5, assigns a display name, and uses ListObject.PutCellValue with row/column offsets (0,0), (0,1), (0,2) to set three header cells before saving the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for the table (e.g., A1:C5)
            int firstRow = 0;      // Row index for A1
            int firstColumn = 0;   // Column index for A1
            int totalRows = 5;     // Total rows including header
            int totalColumns = 3;  // Total columns

            // Add a ListObject (table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);

            ListObject table = sheet.ListObjects[tableIndex];

            // Set a display name for the table (Aspose.Cells uses DisplayName)
            table.DisplayName = "MyTable";

            // Populate header cells using PutCellValue with row and column offsets
            table.PutCellValue(0, 0, "Header1"); // Row offset 0, Column offset 0
            table.PutCellValue(0, 1, "Header2"); // Row offset 0, Column offset 1
            table.PutCellValue(0, 2, "Header3"); // Row offset 0, Column offset 2

            // Determine output path and ensure the directory exists
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.xlsx");
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
