// Title: Create a worksheet table from a cell range and assign a custom display name using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells create a ListObject from a specific cell range with header row | Aspose.Cells set custom display name for a worksheet table in .NET | How to define CellArea and add a table with style in Aspose.Cells C# | Saving a workbook with a named table to XLSX using Aspose.Cells | Apply built‑in table style to a ListObject in Aspose.Cells C#
// Tags: CellArea‑based ListObject generation Aspose.Cells | Assigning DisplayName to Aspose.Cells tables | Setting TableStyleType for Aspose.Cells ListObject | Workbook.Save to XLSX format Aspose.Cells | Renaming worksheet and adding table Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a new Workbook, renames the first worksheet, fills cells A1:C5 with sample data, defines a CellArea covering that range, adds a ListObject (table) with headers, sets its DisplayName to a custom value, applies a built‑in table style, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Populate sample data in the range A1:C5
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

            sheet.Cells["A5"].PutValue(4);
            sheet.Cells["B5"].PutValue("Diana");
            sheet.Cells["C5"].PutValue(88);

            // Define the cell area that will become the table (A1:C5)
            CellArea tableArea = new CellArea
            {
                StartRow = 0,      // Row 1 (zero‑based)
                StartColumn = 0,   // Column A
                EndRow = 4,        // Row 5
                EndColumn = 2      // Column C
            };

            // Add a ListObject (table) to the worksheet using the defined area
            int tableIndex = sheet.ListObjects.Add(
                tableArea.StartRow,
                tableArea.StartColumn,
                tableArea.EndRow,
                tableArea.EndColumn,
                true); // true indicates the first row contains column headers

            ListObject table = sheet.ListObjects[tableIndex];

            // Assign a custom display name to the table (Aspose.Cells uses DisplayName)
            table.DisplayName = "MyCustomTable";

            // Apply a built‑in table style
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Define output path
            string outputPath = "Output.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
