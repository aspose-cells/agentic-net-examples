// Title: Set ShowRowGrandTotals = true on an existing PivotTable using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# enable row grand totals for a pivot table | set ShowRowGrandTotals property on an existing Excel pivot using Aspose | display row totals at the bottom of a pivot report with Aspose.Cells .NET | how to turn on row grand totals in a loaded workbook pivot via C#
// Tags: Aspose.Cells pivot grand total flag | C# modify pivot table settings | Excel workbook pivot configuration Aspose | programmatic pivot table update .NET | Aspose.Cells workbook save after pivot change

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

namespace AsposeCellsExamples
{
    // The example loads an Excel workbook, verifies a PivotTable exists on the first worksheet, sets its ShowRowGrandTotals property to true so row grand totals appear at the bottom of the report, and saves the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Check for the presence of pivot tables
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the first worksheet.");
                }
                else
                {
                    // Retrieve the first pivot table
                    PivotTable pivotTable = worksheet.PivotTables[0];

                    // Enable row grand totals so they appear at the bottom of the report
                    // Use ShowRowGrandTotals which is available across Aspose.Cells versions
                    pivotTable.ShowRowGrandTotals = true;
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
