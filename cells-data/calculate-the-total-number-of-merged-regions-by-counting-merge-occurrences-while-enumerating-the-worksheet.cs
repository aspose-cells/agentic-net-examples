// Title: C# code to count merged regions in an Excel worksheet using Aspose.Cells
// Common Searches: how to count merged cells in an Excel worksheet with Aspose.Cells for .NET | C# example for getting number of merged ranges in a workbook using Aspose.Cells | enumerate merged areas in a worksheet and count them with Aspose.Cells API | Aspose.Cells count total merged regions in used range C# | determine merged region count in an Excel file using Aspose.Cells library
// Tags: Aspose.Cells merged region counting | C# GetMergedRange method usage | worksheet used range iteration Aspose.Cells | Excel merged cells detection .NET | save workbook after processing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file (or creates a new workbook if missing), accesses the first worksheet, iterates through the used cell range, and increments a counter only when a cell is the top‑left cell of a merged region via GetMergedRange, then prints the total merged region count and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load workbook if the file exists; otherwise create a new workbook
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file '{inputPath}' not found. A new workbook will be created.");
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // Get the first worksheet
            Worksheet ws = workbook.Worksheets[0];

            // Counter for merged regions
            int mergedRegionCount = 0;

            // Enumerate all cells in the used range
            for (int row = ws.Cells.MinRow; row <= ws.Cells.MaxRow; row++)
            {
                for (int col = ws.Cells.MinColumn; col <= ws.Cells.MaxColumn; col++)
                {
                    Cell cell = ws.Cells[row, col];

                    // Count only the first cell of each merged region
                    if (cell.IsMerged)
                    {
                        Aspose.Cells.Range mergedRange = cell.GetMergedRange();
                        if (mergedRange != null && mergedRange.FirstRow == row && mergedRange.FirstColumn == col)
                        {
                            mergedRegionCount++;
                        }
                    }
                }
            }

            // Output the result
            Console.WriteLine("Total merged regions: " + mergedRegionCount);

            // Save the workbook (creates the file if it does not exist)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
