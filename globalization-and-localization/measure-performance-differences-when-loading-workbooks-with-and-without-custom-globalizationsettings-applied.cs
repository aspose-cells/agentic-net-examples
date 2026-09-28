// Title: Measure and compare Aspose.Cells workbook load performance with default settings versus custom GlobalizationSettings in C#
// AI Prompts: Create a C# console application that loads an XLSX file with Aspose.Cells, uses Stopwatch to record the load time, then reloads the same file using a LoadOptions instance and sets Workbook.Settings.CultureInfo to a chosen culture, printing both timings. | Enhance the benchmark to loop through several CultureInfo values (e.g., en-US, fr-FR, ja-JP) and output a formatted table showing load duration for each culture. | Add comprehensive error handling that verifies the file path, catches Aspose.Cells exceptions, and logs the worksheet count after each load to confirm successful parsing.
// Common Searches: aspnet measure Aspose.Cells workbook load time with custom culture | how to benchmark Excel file loading speed using LoadOptions in Aspose.Cells | performance impact of setting Workbook.Settings.CultureInfo in Aspose.Cells | compare default and custom globalization settings when loading XLSX with Aspose.Cells | C# timing of Aspose.Cells workbook load with and without LoadOptions
// Tags: benchmark Aspose.Cells workbook loading | custom GlobalizationSettings performance .NET | LoadOptions culture load time measurement | Workbook.Settings.CultureInfo impact on load speed | Excel XLSX load performance Aspose.Cells

using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The program loads 'Sample.xlsx' twice—first with default Aspose.Cells settings and then with a LoadOptions object followed by setting Workbook.Settings.CultureInfo to French—measuring and printing the elapsed milliseconds for each load and displaying the worksheet counts to verify successful loading.
class Program
{
    static void Main()
    {
        // Path to the workbook to be loaded for the test
        string workbookPath = "Sample.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file \"{workbookPath}\" was not found.");
            return;
        }

        try
        {
            // -----------------------------------------------------------------
            // Measure loading time with default settings (no custom globalization)
            // -----------------------------------------------------------------
            Stopwatch swDefault = Stopwatch.StartNew();

            // Load workbook using default settings
            Workbook wbDefault = new Workbook(workbookPath);

            swDefault.Stop();
            Console.WriteLine($"Loading time with default settings: {swDefault.ElapsedMilliseconds} ms");

            // -----------------------------------------------------------------
            // Measure loading time with custom globalization settings applied
            // -----------------------------------------------------------------
            // Create LoadOptions (custom settings can be added here if supported by the version)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            Stopwatch swCustom = Stopwatch.StartNew();

            // Load workbook using the LoadOptions
            Workbook wbCustom = new Workbook(workbookPath, loadOptions);

            // Apply custom culture after loading (affects parsing/formatting)
            wbCustom.Settings.CultureInfo = new CultureInfo("fr-FR");

            swCustom.Stop();
            Console.WriteLine($"Loading time with custom globalization settings: {swCustom.ElapsedMilliseconds} ms");

            // -----------------------------------------------------------------
            // Optional: Verify that both workbooks are loaded correctly
            // -----------------------------------------------------------------
            Console.WriteLine($"Default workbook worksheets count: {wbDefault.Worksheets.Count}");
            Console.WriteLine($"Custom workbook worksheets count: {wbCustom.Worksheets.Count}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
