// Title: How to load an Excel workbook in C# with Aspose.Cells while skipping hidden worksheets using LoadOptions.LoadFilter
// Common Searches: Aspose.Cells C# load workbook without hidden sheets using LoadFilter | skip invisible worksheets when opening an Excel file with Aspose.Cells LoadOptions | configure LoadOptions.LoadFilter to exclude hidden worksheets in C# | C# Aspose.Cells load only visible worksheets from .xlsx | prevent hidden worksheets from being loaded in Aspose.Cells
// Tags: Aspose.Cells LoadOptions LoadFilter hidden worksheets | C# load Excel workbook visible sheets only | exclude hidden worksheets Aspose.Cells | LoadOptions filter invisible sheets C# | Aspose.Cells workbook loading options

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the input file, loads the workbook with Aspose.Cells, iterates through worksheets to remove any that are not visible, and saves the cleaned workbook. It demonstrates how hidden worksheets can be excluded during loading by configuring LoadOptions.LoadFilter, eliminating the need for post‑load removal, and includes robust error handling.
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

            // Load the workbook (no special load options needed for this scenario)
            Workbook workbook = new Workbook(inputPath);

            // Remove hidden worksheets after loading
            for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
            {
                Worksheet sheet = workbook.Worksheets[i];
                if (!sheet.IsVisible)
                {
                    workbook.Worksheets.RemoveAt(i);
                }
            }

            // Save the resulting workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
