// Title: Apply German culture‑specific currency formatting to Excel workbook cells using Aspose.Cells for .NET without altering thread culture
// Common Searches: Aspose.Cells set workbook culture to de-DE for currency formatting only | C# apply German number format to existing Excel file using Aspose.Cells | How to format cells as German Euro currency in Aspose.Cells without changing thread culture | Use built‑in German currency format index 164 in Aspose.Cells .NET example | Iterate through worksheets and apply culture‑specific currency style with Aspose.Cells
// Tags: Workbook.Settings.CultureInfo de-DE Aspose.Cells | apply built‑in German currency format index 164 | culture‑specific number formatting Excel .NET | iterate cells set currency style Aspose.Cells | format currency cells without thread culture change

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, assigns German (de-DE) culture to Workbook.Settings, scans all worksheets to detect currency cells by built‑in format index 164 or currency symbols, reapplies the German built‑in currency format (index 164) to those cells, and saves the updated workbook—all without modifying the thread's culture.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths (replace placeholders with actual paths)
            string inputPath = "{InputFilePath}";
            string outputPath = "{OutputFilePath}";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Apply German culture to the workbook settings (does not affect thread culture)
                workbook.Settings.CultureInfo = new CultureInfo("de-DE");

                // Iterate through all worksheets and cells to adjust currency formatting
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];
                            Style style = cell.GetStyle();

                            bool isCurrency = false;

                            // Check built‑in currency format index (164 is the built‑in German currency format)
                            if (style.Number == 164)
                            {
                                isCurrency = true;
                            }
                            // Check custom format strings that contain typical currency symbols
                            else if (!string.IsNullOrEmpty(style.Custom) &&
                                     (style.Custom.Contains("$") || style.Custom.Contains("€") || style.Custom.Contains("£")))
                            {
                                isCurrency = true;
                            }

                            if (isCurrency)
                            {
                                // Apply the built‑in German currency format (index 164)
                                style.Number = 164;
                                cell.SetStyle(style);
                            }
                        }
                    }
                }

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
