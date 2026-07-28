// Title: Disable chart loading with LoadOptions and capture load warnings in Aspose.Cells for .NET
// Description: Demonstrates how to create a custom LoadFilter that excludes charts, attach a WarningCallback to collect load warnings (e.g., missing chart data), load a workbook with these options, verify that no charts are present, and optionally save the workbook without charts.
// Keywords: Aspose.Cells | C# | LoadOptions | LoadFilter | disable chart loading | warning callback | load warnings | missing chart data | save workbook without charts | memory optimization
// Common Searches: Aspose.Cells load workbook without charts | How to capture load warnings in Aspose.Cells .NET | Exclude charts when opening Excel file using Aspose.Cells | Retrieve warning messages during workbook load | LoadOptions chart flag off Aspose.Cells
// Developer Intent: Open an Excel workbook while preventing chart objects from being loaded and collect any warnings generated about missing chart data.
// Use Cases: Reduce memory usage when processing large workbooks that contain many charts. | Log and analyze warnings about unavailable chart data in automated validation pipelines. | Create a chart‑free copy of a workbook for downstream tasks that only require cell values.
// AI Prompts: Show how to modify the example to also skip images while keeping the warning callback active. | Provide code that filters warnings by type, such as WarningType.ChartDataMissing, and logs them separately. | Explain how to configure LoadOptions to load only formulas and values, excluding charts, images, and other objects.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLoadChartExample
{
    // Custom warning callback to collect load warnings
    // Demonstrates how to create a custom LoadFilter that excludes charts, attach a WarningCallback to collect load warnings (e.g., missing chart data), load a workbook with these options, verify that no charts are present, and optionally save the workbook without charts.
    public class CustomWarningCallback : IWarningCallback
    {
        public List<WarningInfo> Warnings { get; } = new List<WarningInfo>();

        public void Warning(WarningInfo warningInfo)
        {
            // Store each warning for later inspection
            Warnings.Add(warningInfo);
        }
    }

    // Custom load filter that disables loading of charts
    public class CustomLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load everything except charts
            LoadDataFilterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook (should contain charts)
            string inputPath = "input_with_charts.xlsx";
            string outputPath = "output_without_charts.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Please ensure the file exists.");
                return;
            }

            try
            {
                // Prepare load options
                LoadOptions loadOptions = new LoadOptions
                {
                    LoadFilter = new CustomLoadFilter(),                     // Disable chart loading
                    WarningCallback = new CustomWarningCallback()           // Capture warnings
                };

                // Load the workbook with the specified options
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Verify that charts were not loaded
                int chartCount = workbook.Worksheets[0].Charts.Count;
                Console.WriteLine($"Charts loaded in first worksheet: {chartCount}");

                // Retrieve and display load warnings (e.g., missing chart data)
                if (loadOptions.WarningCallback is CustomWarningCallback warningCallback &&
                    warningCallback.Warnings.Count > 0)
                {
                    Console.WriteLine("Load warnings:");
                    foreach (var warning in warningCallback.Warnings)
                    {
                        Console.WriteLine($"- {warning.Description}");
                    }
                }
                else
                {
                    Console.WriteLine("No load warnings were generated.");
                }

                // Optionally save the workbook (charts will be absent)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved without charts to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
