// Title: Load an Excel workbook without charts and capture load warnings using Aspose.Cells for .NET
// Description: Shows how to set LoadOptions with a LoadFilter that omits chart data, attach a custom IWarningCallback to gather WarningInfo objects, open the workbook, and write the warnings to the console. This approach speeds up loading of large files while providing full diagnostic information.
// Keywords: Aspose.Cells | C# | LoadOptions | LoadFilter | exclude charts | load warnings | IWarningCallback | WarningInfo | Excel file loading | performance optimization
// Common Searches: Aspose.Cells load workbook without charts | How to capture load warnings in Aspose.Cells .NET | Disable chart loading with LoadFilter | Implement IWarningCallback for Excel import | LoadOptions example for chart exclusion
// Developer Intent: Open an Excel file while skipping chart objects and retrieve any warnings produced during the load operation.
// Use Cases: Improve load time for massive spreadsheets by ignoring chart data and still monitor compatibility issues. | Validate user‑uploaded Excel files, logging unsupported features while deliberately omitting charts. | Create a centralized warning logger that stores WarningInfo for later reporting or automated quality checks.
// AI Prompts: Modify the example to also exclude images and shapes using LoadFilter. | Write the collected warnings to a JSON file instead of printing them to the console. | Filter the WarningInfo list to show only unsupported formula warnings and summarize them.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Shows how to set LoadOptions with a LoadFilter that omits chart data, attach a custom IWarningCallback to gather WarningInfo objects, open the workbook, and write the warnings to the console. This approach speeds up loading of large files while providing full diagnostic information.
class Program
{
    static void Main()
    {
        // Collect warnings during loading
        List<WarningInfo> warnings = new List<WarningInfo>();
        IWarningCallback warningCallback = new CollectWarningCallback(warnings);

        // Create load options
        LoadOptions loadOptions = new LoadOptions();

        // Disable loading of charts by using a LoadFilter without the Chart flag
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart);

        // Assign the warning callback to capture load warnings
        loadOptions.WarningCallback = warningCallback;

        // Load the workbook with the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Log all captured warnings
        foreach (var warning in warnings)
        {
            Console.WriteLine($"Warning: {warning.Description}");
        }

        // Workbook is now loaded without charts; further processing can be done here
    }

    // Custom implementation of IWarningCallback that stores warnings in a list
    private class CollectWarningCallback : IWarningCallback
    {
        private readonly List<WarningInfo> _warnings;

        public CollectWarningCallback(List<WarningInfo> warnings)
        {
            _warnings = warnings;
        }

        public void Warning(WarningInfo warningInfo)
        {
            _warnings.Add(warningInfo);
        }
    }
}
