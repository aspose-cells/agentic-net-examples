using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomNumberFormatToTsv
{
    class Program
    {
        static void Main()
        {
            // Paths for input and output files
            string sourcePath = "input.xlsx";
            string outputPath = "output.tsv";

            try
            {
                // Verify that the source workbook exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Error: Input file '{sourcePath}' not found.");
                    return;
                }

                // Load the workbook from the file
                Workbook workbook = new Workbook(sourcePath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Create a custom style with a number format (e.g., two decimal places)
                Style customStyle = workbook.CreateStyle();
                customStyle.Custom = "#,##0.00";

                // Apply the custom style to a range of cells (column B rows 2‑5)
                Aspose.Cells.Range range = sheet.Cells.CreateRange("B2:B5");
                range.ApplyStyle(customStyle, new StyleFlag { NumberFormat = true });

                // Prepare text save options for TSV output
                TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    Separator = '\t' // Use tab as the delimiter
                };

                // Save the workbook as a tab‑delimited CSV file
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Workbook saved as tab‑delimited CSV to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}