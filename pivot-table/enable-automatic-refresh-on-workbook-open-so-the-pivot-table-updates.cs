// Title: Refresh all PivotTables in an Excel workbook on open with Aspose.Cells for .NET
// Common Searches: how to refresh all pivot tables in an Excel file using Aspose.Cells C# | Aspose.Cells C# refresh pivot tables before saving workbook | programmatically update pivot cache with Aspose.Cells .NET | auto refresh pivot tables on workbook open using Aspose.Cells | C# code to iterate worksheets and refresh pivot tables Aspose
// Tags: Aspose.Cells PivotTable.RefreshData usage | C# refresh all pivot tables in workbook | update pivot cache before saving Excel | auto refresh pivot tables on workbook open Aspose | iterate worksheets refresh pivot tables C#

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an existing Excel workbook, loops through each worksheet and its PivotTables, calls the RefreshData method to update each pivot cache, and then saves the workbook. This ensures that all pivot tables reflect the latest source data when the file is opened.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Refresh all pivot tables in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    try
                    {
                        // Refresh the pivot table data immediately
                        pivot.RefreshData();
                    }
                    catch (Exception exPivot)
                    {
                        Console.WriteLine($"Failed to refresh pivot table \"{pivot.Name}\": {exPivot.Message}");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
