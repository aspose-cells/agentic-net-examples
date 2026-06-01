using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsNonExponentialDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate numeric data
                sheet.Cells["A1"].PutValue(1234567890L);
                sheet.Cells["A2"].PutValue(0.00000012345);
                sheet.Cells["A3"].PutValue(1.23e+20);

                // Create a style with two decimal places (no scientific notation)
                Style style = workbook.CreateStyle();
                style.Number = 2; // Built‑in format index for two decimal places

                // Apply the style to the range A1:A3
                var range = sheet.Cells.CreateRange("A1:A3");
                range.ApplyStyle(style, new StyleFlag { All = true });

                // Configure TxtSaveOptions to use the display style (formatted values)
                TxtSaveOptions txtOptions = new TxtSaveOptions
                {
                    Separator = '\t', // Tab‑separated values
                    FormatStrategy = CellValueFormatStrategy.DisplayStyle // Use formatted (non‑exponential) values
                };

                // Define output path and ensure the directory exists
                string outputPath = "NonExponentialOutput.txt";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a text file with the specified options
                workbook.Save(outputPath, txtOptions);

                Console.WriteLine("Workbook saved with non‑exponential numeric display.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}