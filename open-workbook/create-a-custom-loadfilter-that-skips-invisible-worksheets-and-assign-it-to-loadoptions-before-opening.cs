// Title: Load an Excel workbook and delete hidden worksheets using Aspose.Cells for .NET
// Common Searches: how to delete hidden worksheets after loading a workbook with Aspose.Cells in C# | Aspose.Cells C# remove invisible sheets from an Excel file | skip hidden worksheets when processing Excel workbooks using Aspose.Cells .NET | programmatically filter out non‑visible sheets with Aspose.Cells load options | C# example to load Excel file and discard hidden worksheets using Aspose.Cells
// Tags: remove hidden worksheets Aspose.Cells | Aspose.Cells sheet visibility handling | C# delete invisible Excel sheets | Aspose.Cells load workbook filter sheets | Excel hidden sheet removal using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLoadFilterExample
{
    // The example loads an Excel file with Aspose.Cells, iterates through its worksheets in reverse order, removes any worksheet whose IsVisible property is false, and saves the resulting workbook to a new file, including file existence validation and exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output files
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook without any filter
                LoadOptions loadOptions = new LoadOptions();
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Remove worksheets that are not visible
                // Iterate backwards to safely delete sheets while looping
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
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
