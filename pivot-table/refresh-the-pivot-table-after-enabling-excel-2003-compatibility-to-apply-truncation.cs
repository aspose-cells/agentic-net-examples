// Title: Refresh all PivotTables in a workbook and save it as Excel 97‑2003 (.xls) using Aspose.Cells for .NET
// Common Searches: how to refresh pivot tables with Aspose.Cells before saving as .xls in C# | Aspose.Cells C# refresh pivot cache and convert workbook to Excel 2003 format | C# iterate worksheets and refresh each PivotTable using Aspose.Cells | save workbook as Excel 97-2003 after updating pivot tables with Aspose.Cells | enable Excel 2003 compatibility and refresh pivot tables in Aspose.Cells .NET
// Tags: pivot table refresh Aspose.Cells | save as Excel 97-2003 Aspose.Cells | recalculate pivot cache C# | xlsx to xls conversion with pivot update | Excel 2003 compatibility truncation Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable class

// The example loads an existing .xlsx workbook, iterates through every worksheet and its PivotTable collection, calls RefreshData and CalculateData on each pivot table to update the cache, then saves the workbook in Excel 97‑2003 (.xls) format, handling output directories and exceptions.
class PivotTableRefresh
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xls";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Refresh all pivot tables in each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pt in sheet.PivotTables)
                {
                    pt.RefreshData();      // Refresh source data
                    pt.CalculateData();   // Recalculate the pivot cache
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook in Excel 97-2003 format
            workbook.Save(outputPath, SaveFormat.Excel97To2003);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
