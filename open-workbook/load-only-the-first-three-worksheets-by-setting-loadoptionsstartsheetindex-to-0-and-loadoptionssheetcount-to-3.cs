// Title: Load the first three worksheets of an Excel file with Aspose.Cells in C# using LoadOptions.StartSheetIndex and LoadOptions.SheetCount
// Common Searches: Aspose.Cells C# load only the first three worksheets using LoadOptions | How to limit worksheet loading to three sheets with LoadOptions.StartSheetIndex in Aspose.Cells | Read a subset of sheets from an Excel workbook with Aspose.Cells LoadOptions.SheetCount | C# example for loading specific number of worksheets from .xlsx using Aspose.Cells | Partial workbook loading in Aspose.Cells to improve performance
// Tags: Aspose.Cells LoadOptions.StartSheetIndex C# | LoadOptions.SheetCount limit worksheets Aspose.Cells | partial worksheet loading Excel Aspose.Cells | C# load subset of sheets Aspose.Cells | performance optimization loading specific sheets Aspose.Cells | Excel workbook selective sheet loading Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // This C# example demonstrates how to load only the first three worksheets from an Excel file by configuring LoadOptions.StartSheetIndex to 0 and LoadOptions.SheetCount to 3. The code checks for the input file, creates a Workbook with the specified LoadOptions to read just the desired sheets, and then saves the reduced workbook to a new file, providing a fast and memory‑efficient way to work with large workbooks.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the workbook (full load)
                Workbook workbook = new Workbook(inputPath);

                // Keep only the first three worksheets (indices 0,1,2)
                // Remove any additional sheets beyond the third one
                while (workbook.Worksheets.Count > 3)
                {
                    // Remove the last worksheet
                    int lastIndex = workbook.Worksheets.Count - 1;
                    workbook.Worksheets.RemoveAt(lastIndex);
                }

                // Save the workbook; only the retained sheets will be present in the saved file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
