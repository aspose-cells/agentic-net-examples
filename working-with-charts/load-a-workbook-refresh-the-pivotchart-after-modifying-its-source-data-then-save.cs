// Title: Modify a worksheet cell, refresh all PivotTables, and automatically update linked PivotCharts in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Change a cell value, call RefreshData and CalculateData on each PivotTable, then save the workbook with Aspose.Cells in C#. | Programmatically ensure that PivotCharts reflect edited source data by recalculating all PivotTables before saving the file using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# recalculate pivot tables after editing source data | Update cell A2 and auto‑refresh linked pivot chart with Aspose.Cells | How to programmatically refresh pivot tables in an Excel workbook using Aspose.Cells | Save workbook after pivot table recalculation using Aspose.Cells .NET example
// Tags: pivot table refresh operation Aspose.Cells | modify worksheet cell Aspose.Cells | linked pivot chart auto update Aspose.Cells | save workbook after pivot refresh .NET | pivot table data update Aspose.Cells Excel

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // PivotTable related classes

// The sample loads 'input.xlsx', changes the value of cell A2 to 12345, iterates through every worksheet to refresh and recalculate each PivotTable (which automatically updates any linked PivotCharts), and then saves the modified workbook as 'output.xlsx'. It includes error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Example modification: change value in cell A2 of the first worksheet
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Cells["A2"].PutValue(12345);

            // Refresh all pivot tables (linked pivot charts are refreshed automatically)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    pivotTable.RefreshData();   // Refresh source data
                    pivotTable.CalculateData(); // Recalculate the pivot table
                }
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
