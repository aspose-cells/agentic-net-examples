// Title: Hide the aggregated values row in an Excel PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide values row in PivotTable | Set ShowValuesRow to false using Aspose.Cells .NET | Remove extra aggregated row from Excel PivotTable programmatically | How to disable the values row in a PivotTable with Aspose.Cells | C# code example for hiding PivotTable values row in Excel
// Tags: Aspose.Cells PivotTable ShowValuesRow property | disable aggregated values row .NET | programmatic PivotTable layout adjustment | Excel workbook save after PivotTable modification | C# Aspose.Cells modify PivotTable layout

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

// The example loads an existing workbook, checks for a PivotTable on the first worksheet, sets the PivotTable's ShowValuesRow property to false to hide the extra aggregated values row, and saves the updated file.
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
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
                return;
            }

            // Access the first PivotTable in the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Hide the additional aggregated values row
            pivotTable.ShowValuesRow = false;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an informative message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
