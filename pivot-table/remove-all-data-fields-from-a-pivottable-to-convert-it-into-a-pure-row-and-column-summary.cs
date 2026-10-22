// Title: How to clear all data fields from the first PivotTable in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# program that loads an .xlsx file with Aspose.Cells, checks the first worksheet for a PivotTable, removes every DataField from that PivotTable, and saves the result to a new file. | Generate Aspose.Cells code that validates the existence of a PivotTable, calls the DataFields.Clear() method to strip all aggregated values, and outputs a cleaned workbook.
// Common Searches: Aspose.Cells C# remove data fields from a pivot table in an existing workbook | How to delete all values from a PivotTable using Aspose.Cells .NET | Clear pivot table data fields programmatically with Aspose.Cells for Excel files | Convert a pivot table to only row and column labels using Aspose.Cells C#
// Tags: Aspose.Cells clear pivot data fields | C# Aspose.Cells remove pivot values | Aspose.Cells pivot table cleanup | Excel pivot table row column summary Aspose | Aspose.Cells DataFields.Clear example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Ensure PivotTable type is available

// The sample loads an Excel workbook, verifies that the first worksheet contains a PivotTable, clears the PivotTable's DataFields collection to leave only row and column headings, and saves the modified workbook to a new file.
class PivotTableCleaner
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            // Get the first pivot table on the sheet
            PivotTable pivotTable = sheet.PivotTables[0];

            // Remove all data fields from the pivot table
            // This converts it into a pure row/column summary without any aggregated values
            pivotTable.DataFields.Clear();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
