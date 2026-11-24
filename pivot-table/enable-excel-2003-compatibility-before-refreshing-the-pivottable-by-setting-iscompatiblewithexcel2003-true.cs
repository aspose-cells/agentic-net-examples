// Title: Set Excel 2003 compatibility for a PivotTable before refreshing it with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set IsCompatibleWithExcel2003 on a pivot table before RefreshData | How to make a pivot table compatible with Excel 2003 using Aspose.Cells and save as .xls | Enable workbook.Settings.IsCompatibleWithExcel2003 in Aspose.Cells before saving to Excel97-2003 format | Refresh pivot table after setting Excel 2003 compatibility flag in Aspose.Cells .NET | Convert .xlsx to .xls with pivot table compatibility using Aspose.Cells C#
// Tags: IsCompatibleWithExcel2003 pivot table Aspose.Cells | refresh pivot data Aspose.Cells C# | save workbook as Excel97To2003 Aspose.Cells | set workbook compatibility Excel 2003 Aspose | convert .xlsx to .xls with pivot compatibility Aspose

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an .xlsx workbook, optionally sets the IsCompatibleWithExcel2003 flag on the workbook and its first pivot table, refreshes and recalculates the pivot data, and then saves the file as an Excel 97‑2003 .xls document using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xls";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Optional: set Excel 2003 compatibility if the property exists in the used version
            // workbook.Settings.IsCompatibleWithExcel2003 = true;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Process pivot tables if present
            if (sheet.PivotTables.Count > 0)
            {
                // Get the first pivot table
                PivotTable pivot = sheet.PivotTables[0];

                // Optional: set compatibility for the pivot table if supported
                // pivot.IsCompatibleWithExcel2003 = true;

                // Refresh and recalculate pivot data
                pivot.RefreshData();
                pivot.CalculateData();
            }

            // Save the workbook in Excel 97-2003 format
            workbook.Save(outputPath, SaveFormat.Excel97To2003);
            Console.WriteLine($"Workbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
