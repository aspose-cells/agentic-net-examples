// Title: Create a macro‑enabled .xlsm workbook with a styled ListObject table using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that creates a new workbook, fills cells A1:C5 with sample data, adds a ListObject table named MyTable, applies a built‑in table style, and saves the file as a macro‑enabled .xlsm workbook. | Show how to programmatically style an Excel table (ListObject) and export the workbook in .xlsm format using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# create macro enabled .xlsm workbook with table | add ListObject table to worksheet using Aspose.Cells .NET | apply built in table style to Excel table with Aspose.Cells | save Aspose.Cells workbook as macro enabled file | populate sample data and create Excel table programmatically Aspose.Cells
// Tags: create macro-enabled workbook Aspose.Cells C# | add ListObject table Aspose.Cells | apply built-in table style Aspose.Cells | save workbook as xlsm Aspose.Cells | populate sample data Aspose.Cells worksheet

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// C# example that uses Aspose.Cells for .NET to build a new workbook, fill cells A1:C5 with sample data, create a ListObject table named MyTable, apply a medium built‑in table style, and save the result as a macro‑enabled .xlsm file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (default format is .xlsx)
            Workbook workbook = new Workbook();

            // Get the first worksheet and set its name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Populate sample data for the table (A1:C5)
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

            // Define the range that will become a table (zero‑based indexes)
            int firstRow = 0;      // A1
            int firstColumn = 0;
            int totalRows = 5;     // including header
            int totalColumns = 3;

            // Add a ListObject (Excel table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "MyTable";               // Set a friendly name
            table.ShowHeaderRow = true;
            table.ShowTableStyleFirstColumn = false;
            table.ShowTableStyleLastColumn = false;
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Save the workbook as a macro‑enabled file (.xlsm)
            string outputPath = "MacroEnabledTable.xlsm";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
