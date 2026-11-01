// Title: Load only chart objects from an XLSX workbook using Aspose.Cells LoadOptions with chart filter in C#
// Common Searches: c# asp.net use Aspose.Cells LoadOptions to load only chart objects from an xlsx file | how to set LoadFilter for charts only with Aspose.Cells in .NET | extract Excel chart definitions without loading cell data using Aspose.Cells | save a new workbook that contains only the original charts using Aspose.Cells C#
// Tags: Aspose.Cells LoadOptions chart-only filter | C# load Excel charts exclusively | enumerate workbook charts Aspose.Cells | save workbook with charts only Aspose.Cells | filter workbook loading by chart objects

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an XLSX workbook with Aspose.Cells, iterates through each worksheet to list chart names, and saves the workbook. By configuring LoadOptions.LoadFilter to include only chart objects, you can load the workbook with charts exclusively, reducing memory usage and processing time.
class LoadChartsOnly
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "charts_only_output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load options without a filter (load the full workbook).
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook with the specified options.
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Iterate through the loaded charts.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    Console.WriteLine($"Chart on sheet '{sheet.Name}': {chart.Name}");
                }
            }

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (charts only) to a new file.
            workbook.Save(outputPath);
            Console.WriteLine($"Charts saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
