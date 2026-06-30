using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Loading;

namespace WorkbookWarningLogger
{
    // Author note: Implements IWarningCallback to capture warnings during workbook load.
    public class WarningCollector : IWarningCallback
    {
        public List<WarningInfo> Warnings { get; } = new List<WarningInfo>();

        public void Warning(WarningInfo warningInfo)
        {
            // Store the warning for later inspection.
            Warnings.Add(warningInfo);
            // Immediate console output for real‑time diagnostics.
            Console.WriteLine($"Warning: {warningInfo.WarningType} - {warningInfo.Description}");
        }
    }

    class Program
    {
        static void Main()
        {
            // Prepare load options with auto‑filter enabled and warning callback attached.
            LoadOptions loadOptions = new LoadOptions
            {
                AutoFilter = true
            };

            WarningCollector collector = new WarningCollector();
            loadOptions.WarningCallback = collector;

            // Load the workbook; warnings (e.g., data loss from filtered rows) will be routed to the collector.
            Workbook workbook = new Workbook("FilteredData.xlsx", loadOptions);

            // Optional: after loading, process collected warnings as needed.
            Console.WriteLine($"Total warnings captured: {collector.Warnings.Count}");

            // Save the workbook if further processing is required.
            workbook.Save("ProcessedWorkbook.xlsx");
        }
    }
}