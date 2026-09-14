// Title: Disable automatic calculation for Excel tables only while keeping other formulas automatic with Aspose.Cells for .NET (C#)
// AI Prompts: Provide a C# example that sets Workbook.Settings.FormulaSettings.CalculationMode to AutomaticExceptTable to stop tables from auto‑recalculating. | Show how to create a ListObject, add a regular formula outside the table, and save the workbook with table‑only calculation disabled using Aspose.Cells. | Generate code that keeps normal worksheet formulas in Automatic mode but excludes ListObject tables from automatic calculation in a .NET workbook.
// Common Searches: Aspose.Cells C# disable auto calculation for ListObject tables only | set calculation mode to AutomaticExceptTable in .NET workbook | keep workbook formulas automatic while turning off table recalculation Aspose.Cells | how to prevent Excel table formulas from auto‑calculating using Aspose.Cells | example of using AutomaticExceptTable mode with Aspose.Cells for .NET
// Tags: AutomaticExceptTable calculation mode Aspose.Cells | disable table auto‑calc .NET | Aspose.Cells ListObject calculation setting | C# workbook calculation mode control | Excel table calculation mode Aspose.Cells | Aspose.Cells formula settings tables

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample creates a workbook, fills cells A1:B3, adds a ListObject (table) over that range, inserts a regular formula in D1, sets Workbook.Settings.FormulaSettings.CalculationMode to AutomaticExceptTable to disable automatic calculation for tables only, and saves the file as TableCalcMode.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some data
            worksheet.Cells["A1"].PutValue(1);
            worksheet.Cells["A2"].PutValue(2);
            worksheet.Cells["A3"].PutValue(3);
            worksheet.Cells["B1"].PutValue(4);
            worksheet.Cells["B2"].PutValue(5);
            worksheet.Cells["B3"].PutValue(6);

            // Define the range for the table (A1:B3)
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int lastRow = firstRow + 2;    // row 3 (zero‑based)
            int lastColumn = firstColumn + 1; // column B (zero‑based)

            // Add the table (ListObject) and set its display name
            int tableIndex = worksheet.ListObjects.Add(firstRow, firstColumn, lastRow, lastColumn, true);
            ListObject table = worksheet.ListObjects[tableIndex];
            table.DisplayName = "SampleTable";

            // Add a regular formula outside the table
            worksheet.Cells["D1"].Formula = "=SUM(A1:B3)";

            // Disable automatic calculation for tables only
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;

            // Determine output file path and ensure the directory exists
            string outputPath = "TableCalcMode.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
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
