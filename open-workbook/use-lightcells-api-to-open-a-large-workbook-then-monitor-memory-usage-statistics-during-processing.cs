// Title: Monitor memory consumption while processing a large XLSX workbook with Aspose.Cells MemoryPreference in C#
// Common Searches: how to monitor GC memory usage while iterating rows of a large XLSX file using Aspose.Cells in C# | C# example for loading large Excel workbook with MemoryPreference to reduce memory footprint | track peak memory during row-by-row processing of a big spreadsheet with Aspose.Cells | retrieve memory statistics with GC.GetTotalMemory while reading Excel rows in .NET | optimize memory when processing large Excel files using Aspose.Cells LoadOptions
// Tags: Aspose.Cells memory preference loading | C# GC memory monitoring Excel processing | large XLSX workbook row iteration memory usage | low‑memory Excel processing with LoadOptions | peak memory tracking Aspose.Cells | Excel row‑by‑row processing low memory

using System;
using System.IO;
using Aspose.Cells;

namespace LightCellsMemoryMonitoring
{
    // The sample loads a large XLSX workbook using Aspose.Cells LoadOptions with MemoryPreference, iterates through each row of the first worksheet, captures current and peak memory via GC.GetTotalMemory, formats the values, and writes memory statistics to the console.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the large workbook file
                string workbookPath = @"C:\Data\LargeWorkbook.xlsx";

                // Verify that the workbook file exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Error: The file \"{workbookPath}\" was not found.");
                    return;
                }

                // Load the workbook (standard loading; LightCells API not required for this example)
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    // Enable memory preference to reduce memory consumption
                    MemorySetting = MemorySetting.MemoryPreference
                };
                Workbook workbook = new Workbook(workbookPath, loadOptions);

                // Select the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Determine the last row with data
                int lastRow = cells.MaxDataRow;

                long peakMemory = 0;

                // Process rows one by one and report memory usage after each row
                for (int row = 0; row <= lastRow; row++)
                {
                    // Example processing: read the entire row (no modification)
                    // This placeholder can be replaced with actual row processing logic
                    var rowCells = cells.Rows[row];

                    // Get current memory usage
                    long currentMemory = GC.GetTotalMemory(forceFullCollection: false);
                    if (currentMemory > peakMemory)
                    {
                        peakMemory = currentMemory;
                    }

                    // Output memory statistics to the console
                    Console.WriteLine($"Row {row + 1}: Current={FormatBytes(currentMemory)} Peak={FormatBytes(peakMemory)}");
                }

                // Final memory usage after processing all rows
                Console.WriteLine("Processing completed.");
                Console.WriteLine($"Final Peak Memory Usage: {FormatBytes(peakMemory)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to format byte values into a readable string
        private static string FormatBytes(long bytes)
        {
            const long KB = 1024;
            const long MB = KB * 1024;
            const long GB = MB * 1024;

            if (bytes >= GB)
                return $"{(double)bytes / GB:0.##} GB";
            if (bytes >= MB)
                return $"{(double)bytes / MB:0.##} MB";
            if (bytes >= KB)
                return $"{(double)bytes / KB:0.##} KB";
            return $"{bytes} B";
        }
    }
}
