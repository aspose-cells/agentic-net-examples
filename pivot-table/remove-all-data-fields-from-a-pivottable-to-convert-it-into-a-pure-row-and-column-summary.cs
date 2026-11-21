// Title: Remove all data fields from an Excel PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to clear data fields in a PivotTable | Delete every data field from an Excel PivotTable programmatically with Aspose.Cells | Convert a PivotTable to a simple row and column summary using Aspose.Cells .NET | Remove pivot table data items without affecting layout in C# Aspose.Cells | Iterate backward to remove PivotTable data fields with Aspose.Cells API
// Tags: Aspose.Cells remove pivot data fields | C# clear Excel PivotTable data collection | Aspose.Cells convert pivot to row column summary | modify PivotTable data fields .NET | programmatic pivot data field deletion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The code loads an Excel workbook, accesses the first worksheet's first PivotTable, removes all entries from its DataFields collection by iterating backwards, and saves the modified workbook to a new file.
class Program
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
            // Load the workbook that contains the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (assumed to contain the PivotTable)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet has at least one PivotTable
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
                return;
            }

            // Get the first PivotTable in the worksheet
            PivotTable pivotTable = sheet.PivotTables[0];

            // Remove all data fields from the PivotTable
            // Iterate backwards because removing items changes the collection size
            for (int i = pivotTable.DataFields.Count - 1; i >= 0; i--)
            {
                pivotTable.DataFields.RemoveAt(i);
            }

            // Save the workbook with the modified PivotTable
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
