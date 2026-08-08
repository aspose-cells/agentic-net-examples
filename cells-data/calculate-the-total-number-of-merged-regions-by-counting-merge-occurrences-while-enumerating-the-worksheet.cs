// Title: How to count merged regions in an Aspose.Cells worksheet using C#
// Description: This example creates a workbook, adds three merged ranges, then scans the used rows and columns. By calling Cell.GetMergedRange() and counting only cells that match the range's FirstRow and FirstColumn, it returns the exact number of distinct merged regions and saves the file.
// Keywords: Aspose.Cells | C# | count merged regions | GetMergedRange | merged cells enumeration | worksheet merged areas | .NET spreadsheet
// Common Searches: count merged cells Aspose.Cells C# | how to get number of merged ranges Aspose.Cells | Aspose.Cells GetMergedRange count example | C# enumerate merged regions in Excel workbook | Aspose.Cells merged area count tutorial
// Developer Intent: Retrieve the total number of distinct merged regions present in a worksheet.
// Use Cases: Validate that a template contains the expected merged‑cell layout before processing. | Generate a summary report of merged region counts for spreadsheet auditing. | Identify merged areas to apply special handling when exporting data to other formats.
// AI Prompts: Write C# code with Aspose.Cells that counts merged regions without double‑counting. | Explain why checking FirstRow and FirstColumn of GetMergedRange prevents duplicate counts. | Show how to list the address of each merged region after counting them.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds three merged ranges, then scans the used rows and columns. By calling Cell.GetMergedRange() and counting only cells that match the range's FirstRow and FirstColumn, it returns the exact number of distinct merged regions and saves the file.
    public class CountMergedRegionsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Add some merged ranges for demonstration
                cells.Merge(0, 0, 2, 2); // A1:B2
                cells.Merge(3, 1, 3, 3); // B4:D6
                cells.Merge(7, 5, 1, 4); // F8:I8

                int mergedRegionCount = 0;

                // Enumerate cells up to the last used row/column
                for (int row = 0; row <= cells.MaxDataRow; row++)
                {
                    for (int col = 0; col <= cells.MaxDataColumn; col++)
                    {
                        Cell cell = cells[row, col];
                        AsposeRange mergedRange = cell.GetMergedRange();

                        // Count only the top‑left cell of each merged area
                        if (mergedRange != null &&
                            mergedRange.FirstRow == row &&
                            mergedRange.FirstColumn == col)
                        {
                            mergedRegionCount++;
                        }
                    }
                }

                Console.WriteLine($"Total number of merged regions: {mergedRegionCount}");

                // Save the workbook (optional)
                workbook.Save("CountMergedRegionsDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
