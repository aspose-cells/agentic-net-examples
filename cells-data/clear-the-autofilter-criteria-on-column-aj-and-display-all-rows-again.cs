using System;
using System.IO;
using Aspose.Cells;

class ClearAutoFilterOnColumnAJ
{
    static void Main()
    {
        const string inputPath = "InputWorkbook.xlsx";
        const string outputPath = "OutputWorkbook.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // If an AutoFilter exists and has any filter columns, clear it
            if (worksheet.AutoFilter != null && worksheet.AutoFilter.FilterColumns.Count > 0)
            {
                // Clear all AutoFilter criteria (including column AJ)
                worksheet.AutoFilter.ShowAll();

                // Refresh the AutoFilter to apply changes (optional but recommended)
                worksheet.AutoFilter.Refresh();
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the filters cleared
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}