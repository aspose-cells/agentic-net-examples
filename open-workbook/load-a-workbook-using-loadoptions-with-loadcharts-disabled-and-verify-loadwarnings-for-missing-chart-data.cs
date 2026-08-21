// Title: Load an Excel workbook without charts using Aspose.Cells LoadOptions and capture missing‑chart warnings via IWarningCallback (C#)
// Description: The example verifies the input file, creates a LoadOptions object that disables chart loading with a LoadFilter, attaches a custom IWarningCallback to receive load warnings, and loads the workbook. Any warnings about missing chart data are reported through the callback, providing a workaround for the unavailable Workbook.LoadWarnings property.
// Keywords: Aspose.Cells LoadOptions | LoadFilter chart exclusion | disable chart loading Aspose.Cells | IWarningCallback C# | load warnings Aspose.Cells | missing chart data warning | C# Excel workbook loading | .NET Aspose.Cells example
// Common Searches: Aspose.Cells load workbook without charts | How to exclude charts when loading an Excel file with Aspose.Cells | Capture load warnings for missing chart data Aspose.Cells | IWarningCallback example for Aspose.Cells | LoadOptions chart filter C#
// Developer Intent: Load a workbook while skipping chart objects and detect any warnings about missing chart data.
// Use Cases: Improve performance by omitting large chart collections during workbook import. | Log or handle warnings when charts are excluded, ensuring data integrity checks. | Replace the deprecated Workbook.LoadWarnings property with a real‑time warning callback.
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, disables chart loading via LoadFilter, and records warnings using a custom IWarningCallback. | Explain step‑by‑step how to use LoadOptions to exclude charts and capture missing‑chart warnings in Aspose.Cells for .NET. | Show how to handle load warnings when Workbook.LoadWarnings is not available, using IWarningCallback in a C# console application.

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input file, creates a LoadOptions object that disables chart loading with a LoadFilter, attaches a custom IWarningCallback to receive load warnings, and loads the workbook. Any warnings about missing chart data are reported through the callback, providing a workaround for the unavailable Workbook.LoadWarnings property.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Create LoadOptions and set a LoadFilter that excludes charts
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFilter = new LoadFilter(LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart),
                WarningCallback = new CustomWarningCallback() // capture warnings via callback
            };

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Since Workbook.LoadWarnings is not available in this version,
            // rely on the warning callback to report any load warnings.
            Console.WriteLine("Workbook loaded successfully (warnings, if any, were reported via callback).");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred while loading the workbook: {ex.Message}");
        }
    }

    // Simple implementation of IWarningCallback to output warnings as they occur
    private class CustomWarningCallback : IWarningCallback
    {
        public void Warning(WarningInfo warningInfo)
        {
            Console.WriteLine($"Callback warning: {warningInfo.Description}");
        }
    }
}
