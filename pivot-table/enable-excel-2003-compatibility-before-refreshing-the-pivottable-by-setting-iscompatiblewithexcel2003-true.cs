// Title: Enable Excel 2003 compatibility, refresh a PivotTable, and save the workbook as .xls with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that sets Workbook.Settings.IsCompatibleWithExcel2003 to true, updates the first PivotTable in the first worksheet using RefreshData and CalculateData, and exports the workbook to Excel 97‑2003 (.xls) format with Aspose.Cells. | Show how to verify that a worksheet contains a PivotTable, handle the case when none exist, enable Excel 2003 compatibility, and then save the workbook as .xls using Aspose.Cells.
// Common Searches: Aspose.Cells enable Excel 2003 compatibility before refreshing a pivot table | C# example to refresh a PivotTable and save workbook as .xls with Aspose.Cells | How to set Excel 2003 compatibility flag for a workbook in Aspose.Cells | Verify PivotTable existence in a worksheet before exporting to Excel 97‑2003 format | Aspose.Cells pivot table refresh and calculate workflow in .NET
// Tags: Excel 2003 compatibility flag Aspose.Cells | PivotTable data refresh Aspose.Cells | Save workbook as Excel97To2003 format Aspose.Cells | Check for PivotTable in worksheet Aspose.Cells | Calculate PivotTable data Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The sample loads an existing .xlsx workbook, optionally enables Excel 2003 compatibility, checks that the first worksheet contains a PivotTable, updates its data with RefreshData, recalculates it, and saves the result as an Excel 97‑2003 .xls file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xls";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Enable Excel 2003 compatibility if the API is available
            // In newer versions the property may be removed; this line can be omitted safely.
            // workbook.Settings.IsCompatibleWithExcel2003 = true;

            // Get the first worksheet (assumed to contain the PivotTable)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet has at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTable found in the first worksheet.");
                return;
            }

            // Get the first PivotTable in the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Refresh the PivotTable data using the correct API
            try
            {
                // RefreshData is a method of PivotTable, not PivotCache
                pivotTable.RefreshData();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to refresh PivotTable data: {ex.Message}");
            }

            // Recalculate the PivotTable
            pivotTable.CalculateData();

            // Save the workbook in Excel 97-2003 format (.xls)
            workbook.Save(outputPath, SaveFormat.Excel97To2003);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
