using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        const string inputPath = "input_with_chart.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Create load options for XLSX files
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Disable loading of charts by using a LoadFilter without the Chart flag
            LoadFilter filter = new LoadFilter(LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart);
            loadOptions.LoadFilter = filter;

            // Set a custom warning callback to capture any warnings during load
            var warningCollector = new WarningCollector();
            loadOptions.WarningCallback = warningCollector;

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Display warnings collected by the callback
            Console.WriteLine("Load warnings:");
            foreach (WarningInfo warning in warningCollector.Warnings)
            {
                Console.WriteLine($"{warning.Type}: {warning.Description}");
            }

            // Save the workbook to confirm successful load
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Custom implementation of IWarningCallback to handle warnings
    private class WarningCollector : IWarningCallback
    {
        public List<WarningInfo> Warnings { get; } = new List<WarningInfo>();

        public void Warning(WarningInfo warningInfo)
        {
            // Store warning information for later processing
            Warnings.Add(warningInfo);
            // Also output warning immediately
            Console.WriteLine($"Callback warning: {warningInfo.Description}");
        }
    }
}