using System;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "input.xlsx";

        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Disable loading of charts by setting a LoadFilter without the Chart option
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart);

        // Set a custom warning callback to collect load warnings
        var warningCollector = new WarningCollector();
        loadOptions.WarningCallback = warningCollector;

        // Load the workbook with the specified options
        Workbook workbook = new Workbook(filePath, loadOptions);

        // Log all collected warnings
        foreach (var warning in warningCollector.Warnings)
        {
            Console.WriteLine($"Warning: {warning.Description}");
        }

        // (Optional) Use the workbook as needed...
    }

    // Custom implementation of IWarningCallback that stores warnings in a list
    class WarningCollector : IWarningCallback
    {
        public List<WarningInfo> Warnings { get; } = new List<WarningInfo>();

        public void Warning(WarningInfo warningInfo)
        {
            Warnings.Add(warningInfo);
        }
    }
}