// Title: Create an Excel table with a dynamic named range that expands automatically using Aspose.Cells for .NET (C#)
// Common Searches: how to add a ListObject with a dynamic named range in Aspose.Cells C# | Aspose.Cells C# create auto expanding Excel table when rows are added | set display name and style for a table using Aspose.Cells .NET | ensure output folder exists before saving workbook with Aspose.Cells C# | save workbook to DynamicTable.xlsx using Aspose.Cells
// Tags: Aspose.Cells create ListObject with dynamic range | C# auto expanding Excel table Aspose.Cells | Aspose.Cells set table display name and style | Aspose.Cells ensure output directory before save | Aspose.Cells save workbook to file .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a new Workbook, adds header and sample rows, defines a ListObject (Excel table) over the range A1:B3, assigns a display name and a built‑in table style, ensures the target directory exists, and saves the file as DynamicTable.xlsx while handling potential exceptions.
class DynamicTableExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");

            // Add some initial data
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Create a table (ListObject) covering the current data range (A1:B3)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
            int tableIndex = sheet.ListObjects.Add(0, 0, 3, 2, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Set a display name for the table
            table.DisplayName = "MyTable";

            // Apply a built‑in table style (optional)
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Determine output file path
            string outputPath = "DynamicTable.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
