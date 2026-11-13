// Title: How to hide column grand totals in an Excel PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide column grand totals in PivotTable | set ShowColumnGrandTotals false using Aspose.Cells | remove column totals from Excel PivotTable programmatically C# | disable column grand totals in PivotTable with Aspose.Cells .NET | Excel PivotTable column grand total suppression using Aspose.Cells API
// Tags: Aspose.Cells PivotTable ShowColumnGrandTotals | C# hide Excel PivotTable column totals | Aspose.Cells modify PivotTable grand totals | Excel PivotTable column grand totals removal .NET | Aspose.Cells column grand totals false | PivotTable column totals suppression C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsRunner
{
    // The example loads an existing workbook, checks for a PivotTable on the first worksheet, sets the PivotTable's ShowColumnGrandTotals property to false to hide column grand totals, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
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

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one PivotTable
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("Error: No PivotTables found in the worksheet.");
                    return;
                }

                // Retrieve the first PivotTable (adjust index if needed)
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Hide the grand totals for columns
                pivotTable.ShowColumnGrandTotals = false;

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display the message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
