// Title: C# – Merge Multiple XLS Workbooks into a Single XLSX with Aspose.Cells
// Description: Loads each legacy .xls file using `new Workbook(filePath)`, validates its presence, combines the workbooks into an empty destination workbook via `Workbook.Combine`, and saves the result as `MergedOutput.xlsx` (SaveFormat.Xlsx) with comprehensive error handling.
// Keywords: Aspose.Cells merge workbooks C# | combine XLS files | Workbook.Combine example | load XLS workbook .NET | convert XLS to XLSX programmatically | C# Excel file consolidation | Aspose.Cells error handling
// Common Searches: how to merge several .xls files into one .xlsx using Aspose.Cells | C# code to combine multiple Excel workbooks | load and merge XLS workbooks Aspose.Cells | Aspose.Cells combine workbooks missing file handling | sample project for merging Excel files in .NET
// Developer Intent: Load each source .xls workbook with `new Workbook(filePath)` and merge them into a single destination workbook.
// Use Cases: Consolidate monthly legacy reports into an annual summary workbook. | Aggregate data from multiple client‑provided XLS files into a master analysis file. | Create a combined workbook from user‑uploaded Excel files in a web service.
// AI Prompts: Write C# code that reads a list of .xls files, checks if each exists, merges them with Aspose.Cells, and saves the result as .xlsx with proper exception handling. | Show an Aspose.Cells example that merges workbooks while preserving original sheet names and resolving duplicates. | Explain the steps to use `Workbook.Combine` and set `SaveFormat.Xlsx` for the output file.

using System;
using System.IO;
using Aspose.Cells;

namespace MergeWorkbooksExample
{
    // Loads each legacy .xls file using `new Workbook(filePath)`, validates its presence, combines the workbooks into an empty destination workbook via `Workbook.Combine`, and saves the result as `MergedOutput.xlsx` (SaveFormat.Xlsx) with comprehensive error handling.
    class Program
    {
        static void Main()
        {
            // Paths of the source XLS workbooks to be merged
            string[] sourceFiles = new string[]
            {
                "Source1.xls",
                "Source2.xls",
                "Source3.xls"
            };

            // Create an empty destination workbook
            Workbook destinationWorkbook = new Workbook();

            foreach (string filePath in sourceFiles)
            {
                try
                {
                    // Verify the source file exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}. Skipping.");
                        continue;
                    }

                    // Load the source workbook
                    Workbook sourceWorkbook = new Workbook(filePath);

                    // Merge the source workbook into the destination workbook
                    destinationWorkbook.Combine(sourceWorkbook);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            try
            {
                // Save the merged workbook to a new file
                destinationWorkbook.Save("MergedOutput.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Merged workbook saved as MergedOutput.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving merged workbook: {ex.Message}");
            }
        }
    }
}
