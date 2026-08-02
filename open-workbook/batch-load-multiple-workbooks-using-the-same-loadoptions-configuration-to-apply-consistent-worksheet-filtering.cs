// Title: Batch load Excel workbooks with shared LoadOptions & custom LoadFilter (C# Aspose.Cells)
// Description: Shows how to build one LoadOptions that holds a CustomLoadFilter—loading full data for visible worksheets and only structure for hidden ones—then loop through multiple .xlsx files, load each workbook with the shared options, output the worksheet count, and save a processed copy with a "_Processed" suffix. Includes file‑existence validation and robust exception handling.
// Keywords: Aspose.Cells | C# | .NET | LoadOptions | LoadFilter | batch workbook processing | Excel file iteration | visible sheet data | hidden sheet structure | performance optimization | error handling | save processed workbook | GitHub example
// Common Searches: Aspose.Cells load multiple workbooks with same LoadOptions | custom LoadFilter for visible and hidden sheets in C# | batch process Excel files using Aspose.Cells | how to apply a shared LoadOptions to several workbooks | reduce hidden sheet size with LoadFilter Aspose.Cells
// Developer Intent: Load a collection of Excel files using one LoadOptions instance that applies a custom LoadFilter to control data loading per worksheet.
// Use Cases: Automate processing of a reports folder, loading full data only for visible sheets while keeping hidden sheets lightweight, then save optimized copies. | Validate worksheet counts across many workbooks after applying a consistent filter to ensure data integrity before analytics. | Create a conversion pipeline that trims hidden sheet content, reducing file size for downstream systems.
// AI Prompts: Generate C# code that scans a directory for .xlsx files and loads each with a shared LoadOptions containing a CustomLoadFilter that loads full data for visible sheets and only structure for hidden sheets, then saves each file with a "_Processed" suffix. | Explain how to extend the CustomLoadFilter to skip formulas in hidden worksheets while preserving cell formatting and comments. | Provide best‑practice error‑logging and continuation logic for batch loading workbooks with Aspose.Cells and a custom LoadFilter.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace BatchLoadExample
{
    // Custom filter that loads all data for visible sheets and only structure for hidden sheets
    // Shows how to build one LoadOptions that holds a CustomLoadFilter—loading full data for visible worksheets and only structure for hidden ones—then loop through multiple .xlsx files, load each workbook with the shared options, output the worksheet count, and save a processed copy with a "_Processed" suffix. Includes file‑existence validation and robust exception handling.
    class CustomLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load full data for visible sheets, only structure for hidden sheets
            LoadDataFilterOptions = sheet.IsVisible
                ? LoadDataFilterOptions.All
                : LoadDataFilterOptions.Structure;
        }
    }

    class Program
    {
        static void Main()
        {
            // Prepare a single LoadOptions instance with the custom filter
            var loadOptions = new LoadOptions
            {
                LoadFilter = new CustomLoadFilter()
            };

            // List of workbook files to be loaded in batch
            var inputFiles = new List<string>
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
            };

            foreach (var filePath in inputFiles)
            {
                try
                {
                    // Verify that the input file exists
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Warning: File not found – skipping '{filePath}'.");
                        continue;
                    }

                    // Load the workbook with the shared LoadOptions
                    var workbook = new Workbook(filePath, loadOptions);

                    // Example operation: output the number of loaded worksheets
                    Console.WriteLine($"File: {filePath} - Loaded Worksheets: {workbook.Worksheets.Count}");

                    // Determine output path
                    var outputPath = Path.Combine(
                        Path.GetDirectoryName(filePath) ?? string.Empty,
                        Path.GetFileNameWithoutExtension(filePath) + "_Processed.xlsx");

                    // Save the processed workbook
                    workbook.Save(outputPath);
                    Console.WriteLine($"Saved processed workbook to: {outputPath}");
                }
                catch (Exception ex)
                {
                    // Catch any runtime exceptions to prevent the program from crashing
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }
        }
    }
}
