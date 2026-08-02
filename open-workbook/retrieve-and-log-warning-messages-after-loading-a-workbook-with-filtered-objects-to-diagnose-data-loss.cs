// Title: Capture Aspose.Cells Load Warnings with IWarningCallback and Structure‑Only LoadFilter (C#)
// Description: Demonstrates how to implement a custom IWarningCallback that records every WarningInfo emitted during workbook loading. The sample configures LoadOptions with a WarningCallback and a LoadFilter that loads only the workbook structure, causing data‑loss warnings. After the Workbook is created, all collected warnings are enumerated and displayed, with an optional save of the partially loaded file.
// Keywords: Aspose.Cells | C# | IWarningCallback | LoadOptions | LoadFilter | Structure filter | warning collection | data loss warning | Excel workbook loading | log warnings
// Common Searches: Aspose.Cells capture load warnings C# | How to use IWarningCallback with LoadFilter | Retrieve warning messages after loading workbook Aspose.Cells | Log data loss warnings Aspose.Cells .NET | Structure load filter warnings Aspose.Cells
// Developer Intent: Collect and log every warning produced when a workbook is loaded with a Structure‑only filter to identify omitted cells, charts, and other elements.
// Use Cases: Audit which workbook components were skipped because of the Structure filter and store the details in a log file. | Display warning summaries in a UI to inform users about potential data loss before saving the workbook. | Programmatically evaluate warnings and decide whether to reload the file with full data based on the presence of DataLoss warnings.
// AI Prompts: Generate a C# snippet that uses Aspose.Cells IWarningCallback to collect load warnings and writes them to a text log. | Show how to filter out WarningType.DataLoss while still capturing other warning types in the WarningCollector. | Provide code that throws an exception if any DataLoss warning is detected after loading a workbook with a Structure filter.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsWarningDemo
{
    // Custom warning callback that stores all warnings for later inspection
    // Demonstrates how to implement a custom IWarningCallback that records every WarningInfo emitted during workbook loading. The sample configures LoadOptions with a WarningCallback and a LoadFilter that loads only the workbook structure, causing data‑loss warnings. After the Workbook is created, all collected warnings are enumerated and displayed, with an optional save of the partially loaded file.
    public class WarningCollector : IWarningCallback
    {
        // List to keep received warnings
        public List<WarningInfo> CollectedWarnings { get; } = new List<WarningInfo>();

        // This method is called by Aspose.Cells whenever a warning occurs
        public void Warning(WarningInfo warningInfo)
        {
            // Store the warning
            CollectedWarnings.Add(warningInfo);

            // Immediate console output (optional)
            Console.WriteLine($"[During Load] Warning: {warningInfo.Type} - {warningInfo.Description}");
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook (replace with actual file)
            string sourceFile = "input.xlsx";

            // Create a warning collector instance
            var warningCollector = new WarningCollector();

            // Configure load options
            LoadOptions loadOptions = new LoadOptions
            {
                // Attach the warning callback
                WarningCallback = warningCollector,

                // Apply a filter that loads only the workbook structure.
                // This will cause data loss for cells, charts, etc., generating warnings.
                LoadFilter = new LoadFilter(LoadDataFilterOptions.Structure)
            };

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // After loading, output all collected warnings
            Console.WriteLine("\n=== Collected Warnings After Loading ===");
            foreach (var warning in warningCollector.CollectedWarnings)
            {
                Console.WriteLine($"Warning Type: {warning.Type}");
                Console.WriteLine($"Description : {warning.Description}");
                Console.WriteLine($"Corrected   : {warning.CorrectedObject ?? "null"}");
                Console.WriteLine(new string('-', 40));
            }

            // (Optional) Save the workbook to verify it is still usable
            workbook.Save("output.xlsx");
        }
    }
}
