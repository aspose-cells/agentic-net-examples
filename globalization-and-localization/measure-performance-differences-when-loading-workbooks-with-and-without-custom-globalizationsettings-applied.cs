// Title: Compare Aspose.Cells workbook load times using the default constructor versus LoadOptions with explicit Xlsx format in C#
// Common Searches: how to measure Aspose.Cells workbook load time in .NET | benchmark Excel file loading speed using Aspose.Cells LoadOptions C# | compare default workbook loading performance to custom LoadOptions in Aspose.Cells | C# code to time Aspose.Cells workbook opening with and without LoadOptions | performance impact of specifying LoadFormat when loading XLSX with Aspose.Cells
// Tags: Aspose.Cells workbook load performance measurement | LoadOptions specifying Xlsx format Aspose.Cells | C# Stopwatch timing Aspose.Cells loading | default constructor vs LoadOptions load speed | Excel file load latency Aspose.Cells .NET

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The sample creates a simple XLSX file if needed, then loads it twice: once with the default Workbook constructor and once with a LoadOptions object that explicitly sets LoadFormat.Xlsx. Stopwatch records the elapsed milliseconds for each load, and the program outputs both timings to illustrate any performance difference between the two loading approaches.
class Program
{
    static void Main()
    {
        // Path to the workbook used for testing
        string filePath = "sample.xlsx";

        // Create a sample workbook if it does not exist
        try
        {
            if (!File.Exists(filePath))
            {
                Workbook wb = new Workbook();
                Worksheet sheet = wb.Worksheets[0];
                for (int i = 0; i < 1000; i++)
                {
                    sheet.Cells[i, 0].PutValue(i);
                    sheet.Cells[i, 1].PutValue(i * 1.1);
                }
                wb.Save(filePath);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating sample workbook: {ex.Message}");
            return;
        }

        // ------------------------------
        // Load without custom LoadOptions
        // ------------------------------
        long defaultLoadMs = 0;
        try
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Workbook file not found.", filePath);

            Stopwatch sw = Stopwatch.StartNew();
            Workbook defaultWb = new Workbook(filePath);
            sw.Stop();
            defaultLoadMs = sw.ElapsedMilliseconds;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook without custom options: {ex.Message}");
        }

        // ------------------------------
        // Load with custom LoadOptions (example: specify format)
        // ------------------------------
        long customLoadMs = 0;
        try
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Workbook file not found.", filePath);

            // Example of using LoadOptions – specify the file format explicitly
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            Stopwatch sw = Stopwatch.StartNew();
            Workbook customWb = new Workbook(filePath, loadOptions);
            sw.Stop();
            customLoadMs = sw.ElapsedMilliseconds;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook with custom options: {ex.Message}");
        }

        // ------------------------------
        // Output the measured times
        // ------------------------------
        Console.WriteLine($"Load time without custom options: {defaultLoadMs} ms");
        Console.WriteLine($"Load time with custom options: {customLoadMs} ms");
    }
}
