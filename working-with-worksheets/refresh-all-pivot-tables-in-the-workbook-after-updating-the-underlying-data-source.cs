// Title: Refresh all PivotTables in an Excel workbook after modifying source data using Aspose.Cells for .NET (C#)
// Common Searches: C# refresh all pivot tables after changing worksheet data with Aspose.Cells | Aspose.Cells .NET update source cells and programmatically refresh pivot tables | Iterate through workbook worksheets to call RefreshData on each PivotTable in C# | Example of using RefreshData method for multiple PivotTables in Aspose.Cells | Save Excel file with refreshed pivot tables after data edit using Aspose.Cells C#
// Tags: Aspose.Cells RefreshData method | refresh all pivot tables .NET | iterate workbook worksheets C# | update source data cells Aspose.Cells | save workbook after pivot refresh

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable class

// The sample loads an existing Excel workbook, updates cells in a data worksheet, then loops through every worksheet and each PivotTable to invoke RefreshData(), ensuring all pivot tables reflect the new data before saving the workbook to a new file.
class PivotTableRefreshExample
{
    static void Main()
    {
        const string inputPath = "InputWorkbook.xlsx";
        const string outputPath = "OutputWorkbook.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // ----- Update underlying data source -----
            // Example: modify some cells that are used by the pivot tables
            Worksheet dataSheet = workbook.Worksheets["Data"]; // assume a sheet named "Data"
            if (dataSheet != null)
            {
                dataSheet.Cells["A2"].PutValue(12345); // sample update
                dataSheet.Cells["B2"].PutValue(67890); // sample update
            }
            else
            {
                Console.WriteLine("Worksheet 'Data' not found.");
            }

            // Refresh all pivot tables in the workbook to reflect the updated data
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pt in sheet.PivotTables)
                {
                    try
                    {
                        pt.RefreshData();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to refresh pivot table '{pt.Name}' on sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with refreshed pivot tables
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
