// Title: Log Workbook Load Warnings with a Custom IWarningCallback and LoadFilter in Aspose.Cells for .NET
// Description: Demonstrates how to attach a custom IWarningCallback to LoadOptions and use a LoadFilter to capture and log warning messages generated while loading an Excel workbook. The collected warnings help developers detect unsupported features or data loss caused by filtered loading.
// Keywords: Aspose.Cells | C# | LoadOptions warning callback | IWarningCallback example | Custom LoadFilter | Excel workbook load warnings | diagnose data loss Aspose.Cells | log warning messages .NET | filtered workbook loading
// Common Searches: Aspose.Cells capture load warnings C# | how to use IWarningCallback with LoadOptions | log warnings when loading Excel with Aspose.Cells | detect data loss using LoadFilter Aspose.Cells | retrieve warning messages after workbook load
// Developer Intent: Capture and log all warning messages produced during a filtered workbook load to identify potential data loss or unsupported features.
// Use Cases: Implement a custom IWarningCallback to collect warnings and write them to a file or monitoring system. | Apply a LoadFilter that loads only specific parts of a workbook (e.g., cell data) while using the warning callback to spot omitted content. | Analyze collected warnings after loading to decide whether to continue processing, alert the user, or retry with different load options.
// AI Prompts: Generate C# code that uses Aspose.Cells LoadOptions with a custom IWarningCallback to log warnings while loading a workbook with a specific LoadFilter. | Show how to filter warning types such as UnsupportedFeature from the callback output and store them in a structured JSON log. | Provide a strategy for handling load warnings: when to abort, when to retry with alternative options, and how to notify end‑users.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsWarningDemo
{
    // Custom warning callback that collects warning messages
    // Demonstrates how to attach a custom IWarningCallback to LoadOptions and use a LoadFilter to capture and log warning messages generated while loading an Excel workbook. The collected warnings help developers detect unsupported features or data loss caused by filtered loading.
    public class CustomWarningCallback : IWarningCallback
    {
        // Store warnings for later inspection
        public List<string> Messages { get; } = new List<string>();

        // This method is invoked by Aspose.Cells when a warning occurs
        public void Warning(WarningInfo warningInfo)
        {
            // Build a readable message
            string msg = $"Warning Type: {warningInfo.Type}, Description: {warningInfo.Description}";
            // Add to collection
            Messages.Add(msg);
            // Also output to console immediately
            Console.WriteLine(msg);
        }
    }

    // Custom load filter to demonstrate filtered loading (e.g., load only cell values and formulas)
    public class CustomLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load only cell data (values, formulas, formatting) for each sheet
            LoadDataFilterOptions = LoadDataFilterOptions.CellData;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook (replace with actual file path)
            string sourceFile = "input.xlsx";

            // Create the custom warning callback instance
            var warningCallback = new CustomWarningCallback();

            // Configure load options
            LoadOptions loadOptions = new LoadOptions
            {
                // Assign the warning callback to capture warnings during loading
                WarningCallback = warningCallback,
                // Apply a custom load filter to limit loaded data
                LoadFilter = new CustomLoadFilter()
            };

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // At this point, any warnings generated during loading have been collected
            Console.WriteLine("\n--- Collected Warning Messages ---");
            if (warningCallback.Messages.Count == 0)
            {
                Console.WriteLine("No warnings were generated during loading.");
            }
            else
            {
                foreach (var msg in warningCallback.Messages)
                {
                    Console.WriteLine(msg);
                }
            }

            // Optional: demonstrate that the workbook is usable after loading
            Console.WriteLine($"\nWorkbook loaded with {workbook.Worksheets.Count} worksheet(s).");
        }
    }
}
