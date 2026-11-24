// Title: Apply a predefined TableStyle to an Aspose.Cells ListObject while preserving original cell formatting in C#
// Common Searches: how to set a built‑in table style for a ListObject in Aspose.Cells C# | preserve cell formatting when applying TableStyleType to an Excel table using Aspose.Cells | Aspose.Cells C# apply TableStyleMedium2 without losing header formatting | add a styled table to a new workbook with Aspose.Cells .NET | C# example for applying predefined table style to a ListObject
// Tags: apply predefined table style ListObject Aspose.Cells | preserve existing cell formatting when styling Excel table | TableStyleMedium2 usage Aspose.Cells C# | create and style Excel table programmatically .NET | listobject styling without formatting loss

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample creates a new workbook, fills it with sample data, defines a range, adds a ListObject (Excel table), assigns a display name, applies the built‑in TableStyleMedium2 style, ensures the output directory exists, and saves the file as StyledTable.xlsx while keeping the original cell formatting intact.
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

            // Populate sample data (including header)
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");

            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["C2"].PutValue(0.5);

            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["C3"].PutValue(0.3);

            sheet.Cells["A4"].PutValue("Orange");
            sheet.Cells["B4"].PutValue(20);
            sheet.Cells["C4"].PutValue(0.6);

            // Define the range for the table (including header row)
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 4;         // header + 3 data rows
            int totalColumns = 3;      // three columns (A‑C)

            // Add a ListObject (table) to the defined range
            int tableIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);

            ListObject table = sheet.ListObjects[tableIndex];

            // Set a display name for the table
            table.DisplayName = "SalesTable";

            // Apply a predefined table style
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Determine output path and ensure directory exists
            string outputPath = "StyledTable.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
