using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsWarningDemo
{
    // Custom warning callback that collects warnings during load/save operations
    public class WarningCollector : IWarningCallback
    {
        // List to store received warnings
        public List<WarningInfo> Warnings { get; } = new List<WarningInfo>();

        // This method is called by Aspose.Cells when a warning occurs
        public void Warning(WarningInfo warningInfo)
        {
            Warnings.Add(warningInfo);
            // Optional: output warning to console for immediate visibility
            Console.WriteLine($"Warning: {warningInfo.WarningType} - {warningInfo.Description}");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook (replace with actual file path)
            string sourcePath = "input.xlsx";

            // Path for the exported XLSX file
            string outputPath = "output.xlsx";

            // Initialize load options and attach the custom warning collector
            LoadOptions loadOptions = new LoadOptions();
            WarningCollector collector = new WarningCollector();
            loadOptions.WarningCallback = collector;

            // Load the workbook using the constructor that accepts a file path and LoadOptions
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // At this point, any processing warnings have been collected in 'collector.Warnings'
            Console.WriteLine($"Total warnings collected during load: {collector.Warnings.Count}");

            // Save the workbook to XLSX format using the Save method with SaveFormat
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // Optionally, display warnings after saving (some operations may generate additional warnings)
            Console.WriteLine($"Total warnings after save: {collector.Warnings.Count}");
        }
    }
}