// Title: C# – Load an Excel workbook without charts and detect missing chart warnings with Aspose.Cells
// Description: Demonstrates how to use Aspose.Cells LoadOptions with a LoadFilter that excludes charts, attach a custom IWarningCallback to capture load warnings, and scan those warnings for messages indicating missing chart data.
// Keywords: Aspose.Cells LoadOptions chart disabled | C# load workbook without charts | Aspose.Cells IWarningCallback example | detect missing chart warnings | LoadFilter exclude charts | Excel chart loading performance
// Common Searches: load excel file without charts asp.net | aspocells capture load warnings c# | exclude charts when opening workbook aspocells | check for missing chart data warning aspocells | custom warning callback aspocells
// Developer Intent: Open a workbook while skipping chart objects and determine if any load warnings report missing chart data.
// Use Cases: Accelerate loading of large spreadsheets by omitting chart rendering. | Validate presence of required chart data without visualizing the charts. | Generate automated reports of chart‑related issues during batch imports.
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, disables chart loading, captures warnings via IWarningCallback, and flags warnings containing "chart". | Explain how to differentiate missing chart data warnings from other chart‑related warnings in Aspose.Cells. | Provide best‑practice guidelines for handling load warnings when charts are excluded in Aspose.Cells for .NET.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Demonstrates how to use Aspose.Cells LoadOptions with a LoadFilter that excludes charts, attach a custom IWarningCallback to capture load warnings, and scan those warnings for messages indicating missing chart data.
class CustomWarningCallback : IWarningCallback
{
    public List<WarningInfo> Warnings { get; } = new List<WarningInfo>();

    public void Warning(WarningInfo warningInfo)
    {
        // Store and display each warning encountered during load
        Warnings.Add(warningInfo);
        Console.WriteLine($"Warning: {warningInfo.Description}");
    }
}

class Program
{
    static void Main()
    {
        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Configure LoadFilter to load everything except charts
        LoadDataFilterOptions filterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart;
        loadOptions.LoadFilter = new LoadFilter(filterOptions);

        // Attach custom warning callback to capture load warnings
        CustomWarningCallback warningCallback = new CustomWarningCallback();
        loadOptions.WarningCallback = warningCallback;

        // Load the workbook with the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Verify if any warning indicates missing chart data
        bool missingChartDataWarning = false;
        foreach (var warning in warningCallback.Warnings)
        {
            // Simple check based on description text
            if (warning.Description != null && warning.Description.IndexOf("chart", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                missingChartDataWarning = true;
                break;
            }
        }

        Console.WriteLine($"Missing chart data warning present: {missingChartDataWarning}");
    }
}
