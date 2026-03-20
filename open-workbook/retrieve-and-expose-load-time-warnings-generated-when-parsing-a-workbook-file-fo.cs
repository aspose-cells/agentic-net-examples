using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWarningDemo
{
    // Custom callback that records all warnings encountered during workbook loading
    public class LoadWarningCollector : IWarningCallback
    {
        // List to store warning information
        public List<WarningInfo> Warnings { get; } = new List<WarningInfo>();

        // This method is called by Aspose.Cells whenever a warning occurs
        public void Warning(WarningInfo warningInfo)
        {
            // Add the warning to the collection for later processing
            Warnings.Add(warningInfo);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Path to the workbook that may generate load-time warnings
            string inputPath = "sample_with_warnings.xlsx";

            // If the sample file does not exist, create a simple workbook for demonstration
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                var ws = tempWb.Worksheets[0];
                ws.Cells["A1"].PutValue("Demo");
                tempWb.Save(inputPath);
            }

            // Create LoadOptions and attach the custom warning collector
            var loadOptions = new LoadOptions
            {
                WarningCallback = new LoadWarningCollector()
            };

            // Load the workbook using the configured options
            var workbook = new Workbook(inputPath, loadOptions);

            // Retrieve the collector from the load options
            var collector = (LoadWarningCollector)loadOptions.WarningCallback;

            // After loading, display all captured warnings
            Console.WriteLine($"Total warnings captured: {collector.Warnings.Count}");
            foreach (var warning in collector.Warnings)
            {
                Console.WriteLine($"Warning Type: {warning.Type}");
                Console.WriteLine($"Description : {warning.Description}");
                Console.WriteLine();
            }

            // Save the workbook to a new file (optional, demonstrates normal save flow)
            workbook.Save("output.xlsx");
        }
    }
}