// Title: Compare memory usage of Automatic and MemoryPreference load modes for large XLSX workbooks with Aspose.Cells in C#
// Common Searches: Aspose.Cells memory usage Automatic vs MemoryPreference load options C# | benchmark memory consumption when loading a 100000 row Excel file with Aspose.Cells | how to measure memory delta of workbook loading using LoadOptions.MemorySetting in .NET | C# code to compare Automatic and MemoryPreference memory modes in Aspose.Cells | performance testing large XLSX workbook load memory footprint Aspose.Cells
// Tags: Aspose.Cells LoadOptions MemorySetting comparison | Automatic load mode memory consumption | MemoryPreference mode memory usage | large XLSX workbook loading performance .NET | GC.GetTotalMemory workbook load measurement

using System;
using System.IO;
using Aspose.Cells;

// The example creates a 100,000‑row by 50‑column XLSX file if missing, then loads it twice—once with the default Automatic memory setting and once with MemoryPreference—while forcing garbage collection and measuring the total memory before and after each load. It outputs the memory delta for each mode, allowing developers to compare the memory impact of the two Aspose.Cells load settings on large workbooks.
class MemoryComparison
{
    static void Main()
    {
        const string filePath = "large.xlsx";

        // Ensure the large workbook exists; create it if necessary
        if (!File.Exists(filePath))
        {
            try
            {
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];

                const int rows = 100_000;
                const int cols = 50;

                // Populate the worksheet with sample data
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        ws.Cells[r, c].PutValue($"R{r}C{c}");
                    }
                }

                wb.Save(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating workbook: {ex.Message}");
                return;
            }
        }

        // Loads a workbook with an optional MemorySetting and returns the memory delta
        long LoadAndMeasureMemory(MemorySetting? setting)
        {
            try
            {
                // Force a full garbage collection before measurement
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                long before = GC.GetTotalMemory(true);

                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

                // Apply the specified memory setting if provided
                if (setting.HasValue)
                {
                    loadOptions.MemorySetting = setting.Value;
                }

                // Load the workbook using the configured options
                Workbook wb = new Workbook(filePath, loadOptions);

                long after = GC.GetTotalMemory(true);
                return after - before;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook{(setting.HasValue ? $" with setting {setting.Value}" : "")}: {ex.Message}");
                return -1;
            }
        }

        // Measure memory consumption using the default (automatic) mode
        long autoMemory = LoadAndMeasureMemory(null);
        if (autoMemory >= 0)
            Console.WriteLine($"Memory used with Automatic mode: {autoMemory / 1024.0 / 1024.0:F2} MB");

        // Measure memory consumption using MemoryPreference mode
        long prefMemory = LoadAndMeasureMemory(MemorySetting.MemoryPreference);
        if (prefMemory >= 0)
            Console.WriteLine($"Memory used with MemoryPreference mode: {prefMemory / 1024.0 / 1024.0:F2} MB");
    }
}
