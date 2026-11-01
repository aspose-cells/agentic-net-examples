// Title: Load an Excel workbook with LoadOptions, keep only sheets 0 and 2, and export to PDF using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# load workbook with LoadOptions and keep specific sheet indices | How to remove all worksheets except selected ones before converting Excel to PDF with Aspose.Cells | C# example for converting only certain Excel sheets to PDF using Aspose.Cells | Load Excel file without charts using LoadOptions in Aspose.Cells C# | Save selected worksheets to PDF programmatically with Aspose.Cells
// Tags: load Excel file using LoadOptions Aspose.Cells | select worksheets by index C# | export selected sheets to PDF Aspose.Cells | remove unwanted worksheets before PDF conversion | disable chart loading LoadOptions Aspose.Cells | programmatic PDF generation from Excel C#

using System;
using System.IO;
using Aspose.Cells;

// The C# program verifies the input Excel file, creates a LoadOptions instance, loads the workbook, removes every worksheet except those at indices 0 and 2, ensures the output directory exists, and saves the remaining sheets as a PDF using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load options – default loading (charts will be loaded; 
            // custom LoadFilter not used due to API version differences)
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Keep only worksheets at index 0 and 2, remove all others
            // Note: after removing a worksheet, indices shift, so iterate backwards
            int[] keepIndices = { 0, 2 };
            for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
            {
                bool keep = false;
                foreach (int idx in keepIndices)
                {
                    if (i == idx)
                    {
                        keep = true;
                        break;
                    }
                }
                if (!keep)
                {
                    workbook.Worksheets.RemoveAt(i);
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the resulting workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
