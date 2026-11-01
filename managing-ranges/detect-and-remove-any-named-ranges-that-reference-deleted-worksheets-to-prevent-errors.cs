// Title: Remove named ranges that reference deleted worksheets using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove named range that refers to a deleted sheet | How to clean up broken named ranges after deleting a worksheet with Aspose.Cells | Detect invalid named ranges in an Excel workbook using Aspose.Cells .NET | Iterate through named ranges and delete those with missing worksheet in Aspose.Cells | C# code to purge named ranges pointing to non‑existent worksheets
// Tags: Aspose.Cells remove invalid named ranges | named range cleanup after worksheet deletion | C# detect broken named range references | Excel workbook range validation Aspose.Cells | iterate named ranges reverse order Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsNamedRangeCleanup
{
    // The example loads an Excel workbook with Aspose.Cells, iterates the named ranges in reverse, attempts to resolve each range, and removes any named range whose range or worksheet is null or throws an exception—indicating a reference to a deleted sheet—before saving the cleaned workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths
            string inputPath = @"C:\Path\To\InputWorkbook.xlsx";
            string outputPath = @"C:\Path\To\CleanedWorkbook.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Get the collection of named ranges in the workbook
            var namedRanges = workbook.Worksheets.Names;

            // Iterate backwards so that removal does not affect the loop index
            for (int i = namedRanges.Count - 1; i >= 0; i--)
            {
                Name name = namedRanges[i];
                bool remove = false;

                try
                {
                    // Attempt to retrieve the range the name refers to
                    Aspose.Cells.Range range = name.GetRange();

                    // If the range or its worksheet is null, the reference is invalid
                    if (range == null || range.Worksheet == null)
                    {
                        remove = true;
                    }
                }
                catch
                {
                    // An exception indicates the reference cannot be resolved (e.g., deleted sheet)
                    remove = true;
                }

                // Remove the named range if it points to a non‑existent worksheet
                if (remove)
                {
                    namedRanges.RemoveAt(i);
                }
            }

            try
            {
                // Save the cleaned workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
