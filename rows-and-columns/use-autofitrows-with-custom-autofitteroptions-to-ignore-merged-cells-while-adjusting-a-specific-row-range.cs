// Title: Auto-fit rows 6‑11 in an Excel sheet while ignoring merged cells with Aspose.Cells for .NET
// AI Prompts: Configure AutoFitterOptions with AutoFitMergedCells set to false and invoke Worksheet.AutoFitRows(startRow, endRow, options) to fit rows 6‑11 without expanding merged cells in C#. | Write a C# routine that loads a workbook, disables merged‑cell auto‑fit via AutoFitterOptions, auto‑fits a specific row range, and saves the updated file.
// Common Searches: Aspose.Cells C# AutoFitRows ignore merged cells for a range of rows | How to prevent merged cells from affecting row auto‑fit in .NET | AutoFitRows with custom AutoFitterOptions example in Aspose.Cells | Fit only rows 6 to 11 in Excel using Aspose.Cells without changing merged cell width | Set AutoFitMergedCells false when auto‑fitting rows in Aspose.Cells
// Tags: AutoFitRows with AutoFitterOptions | ignore merged cells Aspose.Cells | auto‑fit specific row range .NET | disable merged‑cell auto‑fit C# | Excel row height adjustment Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads input.xlsx, creates AutoFitterOptions with AutoFitMergedCells disabled, auto‑fits rows 6‑11 on the first worksheet, and saves the workbook to output.xlsx while handling missing files and exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create AutoFitterOptions and set it to ignore merged cells during auto‑fit
            AutoFitterOptions options = new AutoFitterOptions
            {
                // In versions where IgnoreMergedCells is unavailable, set AutoFitMergedCells to false
                AutoFitMergedCells = false
            };

            // Define the row range to auto‑fit (rows 6 to 11 in Excel are indices 5 to 10)
            int startRow = 5;
            int endRow = 10;

            // Auto‑fit the specified rows using the custom options
            sheet.AutoFitRows(startRow, endRow, options);

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
