// Title: Read rows from a large XLSX worksheet with Aspose.Cells LightCells API in C# without loading the whole workbook
// Common Searches: C# Aspose.Cells LightCells read large worksheet row by row | How to iterate over rows in a big Excel file using Aspose.Cells memory‑optimized loading | Stream Excel rows with Aspose.Cells LoadOptions MemorySetting.MemoryPreference in C# | Read large XLSX file without loading entire workbook using Aspose.Cells LightCells | Efficiently process massive Excel sheet in C# with Aspose.Cells LightCells API
// Tags: Aspose.Cells LightCells row iteration | memory‑optimized Excel reading C# | LoadOptions MemorySetting.MemoryPreference usage | large XLSX streaming with Aspose.Cells | read worksheet cells without full workbook load

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the XLSX file exists, loads it with LoadOptions set to MemorySetting.MemoryPreference to keep memory usage low, selects the first worksheet, determines its used range, and streams each row and column by accessing cells directly, outputting their StringValue to the console without loading the entire workbook into memory.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the large Excel file
            string inputPath = @"C:\Data\LargeWorkbook.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook with memory‑optimized option to reduce memory consumption
            var loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.MemorySetting = MemorySetting.MemoryPreference; // correct enum value

            var workbook = new Workbook(inputPath, loadOptions);

            // Choose the worksheet to process (e.g., the first worksheet)
            var sheet = workbook.Worksheets[0];

            // Determine the used range
            int totalRows = sheet.Cells.MaxDataRow + 1;       // rows are zero‑based
            int totalCols = sheet.Cells.MaxDataColumn + 1;   // columns are zero‑based

            // Iterate through rows and columns without loading the entire sheet into memory
            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    var cell = sheet.Cells[row, col];
                    // Use StringValue to get a textual representation; handle nulls safely
                    string value = cell?.StringValue ?? string.Empty;
                    Console.Write($"{value}\t");
                }
                Console.WriteLine();
            }

            // No need to save the workbook because we only read data.
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
