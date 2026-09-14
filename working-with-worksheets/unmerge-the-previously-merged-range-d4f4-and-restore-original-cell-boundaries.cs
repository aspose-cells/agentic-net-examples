// Title: How to unmerge the D4:F4 range in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, creates a range for cells D4 to F4, calls UnMerge, and saves the workbook. | Show a .NET snippet that accesses the first worksheet, selects the merged range D4:F4, and splits it back into individual cells using the Aspose.Cells API. | Write a C# program that verifies an input Excel file, uses Aspose.Cells to unmerge the cells D4 through F4, and writes the result to a new file.
// Common Searches: Aspose.Cells C# unmerge cells D4-F4 in existing workbook | How to split merged range D4 to F4 using Aspose.Cells .NET | C# code to remove merged cells from an Excel file with Aspose.Cells | Unmerge specific cell range in .xlsx using Aspose.Cells library
// Tags: Aspose.Cells UnMerge method usage | C# Excel range unmerge | programmatic Excel cell split .NET | modify merged cells with Aspose.Cells | unmerge specific Excel range .NET

using System;
using System.IO;
using Aspose.Cells;

// // Loads input.xlsx, creates a range covering D4:F4, calls UnMerge() to separate the merged cells, and saves the result to output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains the merged range D4:F4
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // D4:F4 corresponds to row 3 (zero‑based) and columns 3 to 5
            int firstRow = 3;      // Row D4 (zero‑based)
            int firstColumn = 3;   // Column D (zero‑based)
            int totalRows = 1;     // Only one row in the merged range
            int totalColumns = 3;  // Columns D, E, F

            // Create a range representing the merged cells and unmerge it
            Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns);
            mergedRange.UnMerge();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
