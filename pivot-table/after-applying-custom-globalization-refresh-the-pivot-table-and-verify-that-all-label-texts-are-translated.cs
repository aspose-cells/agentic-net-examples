// Title: Refresh all PivotTables in an Excel workbook after applying custom globalization using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells refresh pivot tables after changing workbook culture | How to recalculate pivot table data programmatically with Aspose.Cells .NET | Iterate through worksheets and refresh every pivot table in an Excel file using Aspose.Cells | Save workbook after pivot table refresh with Aspose.Cells example in C# | Handle exceptions while refreshing pivot tables in Aspose.Cells
// Tags: Aspose.Cells refresh pivot tables | Aspose.Cells calculate pivot data | C# iterate worksheets pivot tables | save workbook after pivot refresh Aspose.Cells | error handling pivot table processing C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable class

// The sample loads an existing Excel workbook, loops through each worksheet and its PivotTables, calls RefreshData and CalculateData to update source data and recalculate results, catches any processing errors, and then saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Refresh and recalculate all pivot tables in the workbook
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (PivotTable pt in ws.PivotTables)
                {
                    try
                    {
                        pt.RefreshData();      // Refresh source data
                        pt.CalculateData();   // Recalculate the pivot table
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to process PivotTable '{pt.Name}' in sheet '{ws.Name}': {ex.Message}");
                    }
                }
            }

            Console.WriteLine("All pivot tables have been refreshed and recalculated.");

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }
}
