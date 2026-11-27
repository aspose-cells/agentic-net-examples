// Title: C# example: Load only the first three worksheets from an XLSX file with Aspose.Cells while keeping memory usage low
// Common Searches: aspnet load first three sheets from xlsx using Aspose.Cells without loading all worksheets | c# reduce memory consumption when opening large Excel file with Aspose.Cells | how to keep only selected worksheets after opening workbook in Aspose.Cells C# | Aspose.Cells load options to limit worksheets loaded in C# | remove extra worksheets programmatically after loading workbook with Aspose.Cells
// Tags: load first three worksheets Aspose.Cells C# | memory efficient Excel loading Aspose.Cells | selective worksheet loading LoadOptions Aspose.Cells | remove worksheets after workbook load Aspose.Cells | limit loaded sheets Aspose.Cells Xlsx

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The C# program verifies the source XLSX file, loads it with Aspose.Cells, removes all worksheets beyond the third one to retain only the first three sheets, and saves the trimmed workbook, providing a memory‑efficient way to work with large Excel files.
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file.
            string sourcePath = "input.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException.
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Load the workbook without a custom filter (all worksheets are loaded).
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
                Workbook workbook = new Workbook(sourcePath, loadOptions);

                // Keep only the first three worksheets.
                for (int i = workbook.Worksheets.Count - 1; i >= 3; i--)
                {
                    workbook.Worksheets.RemoveAt(i);
                }

                // Verify the number of loaded worksheets.
                Console.WriteLine($"Worksheets retained: {workbook.Worksheets.Count}");

                // Save the workbook to a new file (contains only the first three sheets).
                string outputPath = "output_first_three_sheets.xlsx";

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display an error message.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
