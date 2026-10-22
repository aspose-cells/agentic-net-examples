// Title: How to enable PreserveFormatting on an Aspose.Cells PivotTable in C# to retain styles after refresh
// AI Prompts: Load an Excel workbook with Aspose.Cells, locate the first PivotTable, set its PreserveFormatting property to true, and save the workbook. | Write C# code that iterates through all PivotTables in a worksheet and enables PreserveFormatting to keep formatting after each refresh. | Show how to verify a PivotTable exists before applying PreserveFormatting using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set PreserveFormatting on PivotTable to keep cell styles after refresh | How to retain pivot table formatting when refreshing data with Aspose.Cells .NET | C# example for enabling PreserveFormatting flag on Excel PivotTable using Aspose.Cells | Preserve formatting of PivotTable after refresh in Aspose.Cells workbook
// Tags: Aspose.Cells PivotTable PreserveFormatting property | C# enable formatting retention for Excel PivotTable | set PreserveFormatting flag in Aspose.Cells workbook | retain pivot table styles after refresh .NET | modify PivotTable formatting behavior Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExample
{
    // The program loads an existing Excel workbook, accesses the first worksheet's first PivotTable, sets its PreserveFormatting property to true so that applied styles persist after each refresh, and saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (assumed to contain the PivotTable)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one PivotTable on the worksheet
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found on the first worksheet.");
                    return;
                }

                // Retrieve the first PivotTable
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Preserve formatting after each refresh
                pivotTable.PreserveFormatting = true;

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
