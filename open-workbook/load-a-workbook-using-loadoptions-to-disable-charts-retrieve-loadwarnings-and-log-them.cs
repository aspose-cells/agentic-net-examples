// Title: Aspose.Cells .NET – Load workbook without charts and capture load warnings
// Description: Demonstrates how to configure LoadOptions with a LoadFilter that excludes chart data, attach a custom IWarningCallback to collect WarningInfo objects, open an Excel file, and log each warning’s type and description.
// Keywords: Aspose.Cells | .NET | LoadOptions | LoadFilter | disable charts | load warnings | IWarningCallback | Excel workbook loading | performance optimization | GitHub example
// Common Searches: Aspose.Cells load workbook without charts | capture load warnings Aspose.Cells .NET | how to use LoadFilter to skip charts | custom warning callback Aspose.Cells | improve Excel load performance Aspose
// Developer Intent: Open an Excel file while skipping chart loading and retrieve any load‑time warnings.
// Use Cases: Speed up processing of large workbooks by omitting chart data. | Audit unsupported or altered features after import via collected warnings. | Trigger custom remediation when specific warning types (e.g., unknown chart types) are detected. | Log warning details to external systems for compliance reporting.
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, disables chart loading, and writes all load warnings to a CSV file. | Show how to implement an IWarningCallback that stores WarningInfo records in a SQL database after loading a workbook with a custom LoadFilter. | Create a script that uses LoadOptions to exclude charts, processes the collected WarningInfo objects, and generates a markdown summary report.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsLoadOptionsDemo
{
    // Custom warning callback that stores warnings in a list
    // Demonstrates how to configure LoadOptions with a LoadFilter that excludes chart data, attach a custom IWarningCallback to collect WarningInfo objects, open an Excel file, and log each warning’s type and description.
    public class CollectingWarningCallback : IWarningCallback
    {
        public List<WarningInfo> Warnings { get; } = new List<WarningInfo>();

        public void Warning(WarningInfo warningInfo)
        {
            // Store each warning for later processing
            Warnings.Add(warningInfo);
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string inputPath = "input.xlsx";

            // Create a warning collector
            var warningCallback = new CollectingWarningCallback();

            // Create LoadOptions and configure it:
            // - Disable loading of charts by excluding the Chart flag
            // - Attach the warning callback to capture load warnings
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.WarningCallback = warningCallback;
            loadOptions.LoadFilter = new LoadFilter(
                LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart);

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Log all captured warnings
            Console.WriteLine("Load warnings:");
            foreach (var warning in warningCallback.Warnings)
            {
                Console.WriteLine($"- Type: {warning.WarningType}, Description: {warning.Description}");
            }

            // (Optional) Use the workbook as needed...
            // For demonstration, output the number of worksheets loaded
            Console.WriteLine($"Worksheets loaded: {workbook.Worksheets.Count}");
        }
    }
}
