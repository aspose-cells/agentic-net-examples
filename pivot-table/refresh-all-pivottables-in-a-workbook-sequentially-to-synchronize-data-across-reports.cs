// Title: Sequentially refresh every PivotTable in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens a workbook, validates the input file, iterates through each worksheet and its PivotTables, calls RefreshData on each, and saves the result to a new file with Aspose.Cells. | Add comprehensive try‑catch error handling and logging to a routine that refreshes all PivotTables in a workbook using Aspose.Cells. | Extend the example to refresh all PivotTables and then export the workbook to PDF in a single workflow with Aspose.Cells.
// Common Searches: C# Aspose.Cells loop through worksheets to refresh all pivot tables | Refresh PivotTable data source before saving workbook using Aspose.Cells .NET | Check if Excel file exists then refresh pivot tables with Aspose.Cells | Batch refresh of multiple pivot tables in a workbook using Aspose.Cells C# | How to call PivotTable.RefreshData for every table in an Excel file with Aspose.Cells
// Tags: refresh pivot tables Aspose.Cells | iterate worksheets pivot tables C# | PivotTable RefreshData Aspose.Cells | save workbook after pivot refresh | file existence validation Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The sample loads an existing Excel workbook, verifies the source file, iterates over each worksheet and each PivotTable to invoke RefreshData, and then saves the updated workbook to a new file while handling exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook containing the PivotTables
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each PivotTable on the current worksheet
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    // Refresh the data source of the PivotTable
                    pivotTable.RefreshData();
                }
            }

            // Save the workbook after all PivotTables have been refreshed
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
