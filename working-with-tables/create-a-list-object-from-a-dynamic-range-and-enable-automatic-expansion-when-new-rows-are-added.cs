// Title: Create a dynamic ListObject (Excel table) from the worksheet's used range that auto‑expands when new rows are added using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that determines the worksheet's used range, creates a ListObject over that range, applies a table style, and ensures the table grows automatically as rows are appended. | Show how to save the workbook after adding a dynamically sized ListObject and verify that the table expands without an explicit AutoExpand setting.
// Common Searches: aspnet create Excel table from used range that expands automatically Aspose.Cells | c# Aspose.Cells ListObject auto expand when adding rows | how to make an Excel table dynamic with Aspose.Cells .NET | Aspose.Cells add styled ListObject based on MaxDataRow MaxDataColumn | save workbook after creating dynamic table Aspose.Cells C#
// Tags: create ListObject from used range Aspose.Cells | dynamic Excel table without AutoExpand property | apply TableStyleMedium9 to ListObject | save workbook as .xlsx using Aspose.Cells | calculate used range with MaxDataRow MaxDataColumn

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The program creates a new workbook, writes header and sample data, computes the used range, adds a styled ListObject (Excel table) over that range, and saves the file; the table automatically expands when additional rows are inserted within the worksheet.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            // Populate sample data
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(90);

            // Determine the used range (including header)
            int firstRow = 0;               // zero‑based index for row 1
            int firstColumn = 0;            // zero‑based index for column A
            int totalRows = sheet.Cells.MaxDataRow + 1;
            int totalColumns = sheet.Cells.MaxDataColumn + 1;

            // Add a ListObject (Excel table) based on the current range
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows - 1, totalColumns - 1, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Set table name and style
            table.DisplayName = "DynamicTable";
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Note: Aspose.Cells does not provide an AutoExpand property.
            // The table will expand automatically when new rows are added within the worksheet range.

            // Define output file path
            string outputPath = "DynamicListObject.xlsx";

            // Ensure the output directory exists
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
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
