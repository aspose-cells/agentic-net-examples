// Title: Refresh all pivot tables in an Excel workbook after disabling Excel 2003 compatibility with Aspose.Cells for .NET
// Common Searches: how to programmatically refresh pivot tables in a .xlsx file using Aspose.Cells C# | Aspose.Cells C# example to recalculate pivot table data after turning off Excel 2003 compatibility mode | iterate through worksheets and refresh each pivot table with Aspose.Cells .NET | C# code to call RefreshData and CalculateData on Aspose.Cells PivotTable objects | save workbook after updating pivot tables using Aspose.Cells for .NET | handling pivot table refresh errors with Aspose.Cells C#
// Tags: Aspose.Cells RefreshData pivot table | Aspose.Cells CalculateData method | C# refresh all pivot tables in workbook | disable Excel 2003 compatibility Aspose.Cells | save workbook after pivot refresh Aspose.Cells | error handling for Aspose.Cells pivot tables

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable class

// The sample loads an existing .xlsx workbook, iterates over each worksheet and its PivotTable collection, calls RefreshData to update the source data and CalculateData to recompute values, handles any pivot‑specific exceptions, ensures the output directory exists, and saves the modified workbook to a new file.
class PivotTableRefresh
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Refresh all pivot tables in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each pivot table on the worksheet
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    try
                    {
                        // Refresh the data source of the pivot table
                        pivot.RefreshData();

                        // Recalculate the pivot table values
                        pivot.CalculateData();
                    }
                    catch (Exception exPivot)
                    {
                        Console.WriteLine($"Error refreshing pivot table '{pivot.Name}' on sheet '{sheet.Name}': {exPivot.Message}");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
