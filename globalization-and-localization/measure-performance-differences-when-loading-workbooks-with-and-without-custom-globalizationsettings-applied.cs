// Title: Measure Aspose.Cells workbook loading performance with and without custom GlobalizationSettings in C#
// Common Searches: C# how to benchmark Aspose.Cells workbook load time using Stopwatch | compare load speed of Excel file in Aspose.Cells when applying custom CultureInfo | measure impact of GlobalizationSettings on Aspose.Cells workbook opening time | performance test default vs custom number format culture in Aspose.Cells | Aspose.Cells load time difference after setting Settings.CultureInfo in C#
// Tags: Aspose.Cells workbook load performance | custom GlobalizationSettings load latency | C# Stopwatch timing Aspose.Cells | apply CultureInfo to workbook Settings | Excel loading speed with custom number format

using System;
using System.Diagnostics;
using System.IO;
using System.Globalization;
using Aspose.Cells;

// The example creates a sample XLSX file, then loads it twice: first with the default settings and second after assigning a cloned invariant CultureInfo with modified decimal, group, and currency symbols to Workbook.Settings.CultureInfo. Each load is timed with Stopwatch, and the elapsed milliseconds are printed to compare the performance impact of custom GlobalizationSettings.
class Program
{
    static void Main()
    {
        // Path to the workbook used for testing
        string filePath = "sample.xlsx";

        try
        {
            // Create a sample workbook if it does not exist
            if (!File.Exists(filePath))
            {
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                ws.Name = "Data";

                // Populate with data to make loading non‑trivial
                for (int row = 0; row < 1000; row++)
                {
                    for (int col = 0; col < 50; col++)
                    {
                        ws.Cells[row, col].PutValue(row * col);
                    }
                }

                wb.Save(filePath);
            }

            // ------------------------------
            // Load with default settings
            // ------------------------------
            Stopwatch swDefault = Stopwatch.StartNew();
            Workbook wbDefault = new Workbook(filePath);
            swDefault.Stop();

            // ------------------------------
            // Load and apply custom globalization settings
            // ------------------------------
            Stopwatch swCustom = Stopwatch.StartNew();
            Workbook wbCustom = new Workbook(filePath);

            // Clone the invariant culture and modify its number format settings
            CultureInfo customCulture = (CultureInfo)CultureInfo.InvariantCulture.Clone();
            customCulture.NumberFormat.NumberDecimalSeparator = ",";
            customCulture.NumberFormat.NumberGroupSeparator = ".";
            customCulture.NumberFormat.CurrencySymbol = "€";

            // Apply the custom culture to the workbook
            wbCustom.Settings.CultureInfo = customCulture;

            swCustom.Stop();

            // ------------------------------
            // Output the measured times
            // ------------------------------
            Console.WriteLine($"Load time with default settings: {swDefault.ElapsedMilliseconds} ms");
            Console.WriteLine($"Load time with custom globalization settings: {swCustom.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
