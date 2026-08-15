// Title: C# – Unmerge cells C6:E7, recalculate formulas, and save workbook using Aspose.Cells
// Description: Loads an existing Excel file, accesses the first worksheet, unmerges the merged range C6:E7, triggers a full formula recalculation, and saves the updated workbook. Includes file‑existence check and exception handling for robust automation.
// Keywords: Aspose.Cells unmerge cells C6:E7 | C# calculate formulas Aspose.Cells | Excel unmerge range .NET | Workbook.CalculateFormula example | Aspose.Cells load and save workbook | unmerge merged cells programmatically | Excel automation C# Aspose
// Common Searches: how to unmerge a specific range with Aspose.Cells C# | Aspose.Cells recalculate all formulas after unmerge | C# code to split merged cells C6:E7 in Excel | Aspose.Cells unmerge cells and save workbook | calculate formulas after modifying worksheet Aspose
// Developer Intent: Load an existing Excel workbook, unmerge the merged range C6:E7, recalculate every formula, and write the modified file to a new location.
// Use Cases: Prepare a template for data import by removing merged cells that break parsing. | Ensure accurate calculations in generated reports after programmatically changing layout. | Clean up workbooks received from external sources where merged cells cause formula errors.
// AI Prompts: Write C# code with Aspose.Cells that checks for an input file, unmerges range C6:E7, runs CalculateFormula, and saves to a new file with error handling. | Show how to safely call UnMerge() on a range that may already be unmerged, and then recalculate all formulas in the workbook. | Explain best practices for loading, modifying (unmerge), recalculating, and saving Excel files using Aspose.Cells in .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsUnmergeAndRecalculate
{
    // Loads an existing Excel file, accesses the first worksheet, unmerges the merged range C6:E7, triggers a full formula recalculation, and saves the updated workbook. Includes file‑existence check and exception handling for robust automation.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (index 0)
                Worksheet worksheet = workbook.Worksheets[0];

                // Create a range representing the merged cells C6:E7 and unmerge them
                Aspose.Cells.Range mergedRange = worksheet.Cells.CreateRange("C6", "E7");
                mergedRange.UnMerge();

                // Recalculate all formulas in the workbook
                workbook.CalculateFormula();

                // Save the modified workbook to the output file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
