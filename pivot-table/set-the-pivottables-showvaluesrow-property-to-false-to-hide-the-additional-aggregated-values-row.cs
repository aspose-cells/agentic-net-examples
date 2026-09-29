// Title: How to hide the aggregated values row of a PivotTable in an Excel file with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an existing .xlsx workbook, accesses the first PivotTable, sets ShowValuesRow to false, and saves the modified file using Aspose.Cells. | Write a C# snippet that verifies a PivotTable exists on the first worksheet, disables its values row display, and writes the updated workbook to a new location with Aspose.Cells. | Provide a step‑by‑step C# example showing how to hide the extra values row in a PivotTable by configuring ShowValuesRow = false via Aspose.Cells.
// Common Searches: aspnet hide values row in pivot table using Aspose.Cells C# example | set ShowValuesRow false Aspose.Cells .NET tutorial | remove aggregated values row from Excel pivot table programmatically C# | Aspose.Cells hide pivot table values row without recreating pivot | C# code to disable values row in existing Excel pivot table Aspose
// Tags: Aspose.Cells ShowValuesRow property | hide pivot values row C# | modify existing pivot table Aspose.Cells | Excel pivot table values row removal .NET | C# Aspose.Cells workbook pivot customization

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// // Loads input.xlsx, checks for a PivotTable on the first worksheet, sets its ShowValuesRow property to false to hide the aggregated values row, and saves the result as output.xlsx.
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
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (assumed to contain the PivotTable)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTable found in the first worksheet.");
                return;
            }

            // Retrieve the first PivotTable
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Hide the additional aggregated values row
            pivotTable.ShowValuesRow = false;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
