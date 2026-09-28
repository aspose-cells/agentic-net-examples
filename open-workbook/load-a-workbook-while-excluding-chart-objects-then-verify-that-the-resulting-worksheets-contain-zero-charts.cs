// Title: Load an Excel workbook with Aspose.Cells for .NET, delete all charts from every worksheet, and confirm zero charts remain
// AI Prompts: Generate C# code that opens an .xlsx file using Aspose.Cells, clears the Charts collection on each worksheet, and prints a message indicating whether any charts are left. | Write a method that takes a file path, loads the workbook with Aspose.Cells, removes all chart objects from all worksheets, and returns true only if the workbook contains no charts after the operation.
// Common Searches: Aspose.Cells C# remove charts from workbook and check chart count | How to clear chart collections on each worksheet using Aspose.Cells .NET | Verify that an Excel file loaded with Aspose.Cells has zero charts | C# load Excel workbook without loading chart objects Aspose.Cells | Aspose.Cells example to delete all charts and validate removal
// Tags: Aspose.Cells clear worksheet charts C# | Aspose.Cells verify zero charts .NET | load Excel workbook without chart objects Aspose.Cells | remove chart collections from Excel worksheets C# | chart count validation using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Loads an Excel file with Aspose.Cells, clears all chart objects from each worksheet, and verifies that every worksheet ends up with zero charts, outputting the verification result.
class Program
{
    static void Main()
    {
        // Path to the source workbook
        string sourcePath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Error: The file \"{sourcePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook (default LoadOptions)
            Workbook workbook = new Workbook(sourcePath);

            // Remove all chart objects from each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Charts.Count > 0)
                {
                    sheet.Charts.Clear();
                }
            }

            // Verify that each worksheet now contains zero charts
            bool allSheetsHaveNoCharts = true;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Charts.Count > 0)
                {
                    allSheetsHaveNoCharts = false;
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" still contains {sheet.Charts.Count} chart(s).");
                }
            }

            if (allSheetsHaveNoCharts)
            {
                Console.WriteLine("Verification passed: all worksheets contain zero charts.");
            }
            else
            {
                Console.WriteLine("Verification failed: some worksheets still contain charts.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
