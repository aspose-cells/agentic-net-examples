// Title: Remove a pivot table from an Excel worksheet and confirm its deletion using Aspose.Cells for .NET (C#)
// Common Searches: C# code to delete a pivot table in an Excel file with Aspose.Cells and verify count is zero | How to programmatically remove the first pivot table from a worksheet using Aspose.Cells for .NET | Aspose.Cells check if worksheet has any pivot tables after removal in C# | Validate that no pivot tables remain after calling Worksheet.PivotTables.RemoveAt in C# | Aspose.Cells example for deleting pivot tables and confirming empty PivotTables collection
// Tags: Aspose.Cells delete pivot table C# | verify empty PivotTables collection Aspose.Cells | remove first pivot table worksheet Aspose.Cells | pivot table count after removal Aspose.Cells | C# Excel pivot table deletion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, accesses the first worksheet, removes the first pivot table if present, checks that the worksheet's PivotTables collection is empty, outputs the verification result, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine input and output file paths (use defaults if not provided)
            string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
            string outputPath = args.Length > 1 ? args[1] : "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
                return;
            }

            try
            {
                // Load the workbook from the specified input file
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (index 0)
                Worksheet worksheet = workbook.Worksheets[0];

                // Delete the first pivot table if any exist
                if (worksheet.PivotTables.Count > 0)
                {
                    // Remove the pivot table from the collection
                    worksheet.PivotTables.RemoveAt(0);
                }

                // Verify that the worksheet no longer contains any pivot tables
                bool hasNoPivotTables = worksheet.PivotTables.Count == 0;

                // Output verification result
                Console.WriteLine("Pivot tables remaining: " + worksheet.PivotTables.Count);
                Console.WriteLine("Verification passed: " + hasNoPivotTables);

                // Save the modified workbook to the specified output file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display an error message
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
