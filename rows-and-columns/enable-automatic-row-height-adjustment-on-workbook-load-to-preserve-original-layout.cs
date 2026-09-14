// Title: Automatically auto‑fit all row heights when loading an XLSX workbook with Aspose.Cells for .NET
// AI Prompts: Load an existing XLSX file with Aspose.Cells using LoadOptions that include AutoFitterOptions (OnlyAuto = false) to auto‑fit every row height, then save the workbook. | Configure AutoFitterOptions to ignore custom row heights and apply auto‑fit on load, demonstrating how to preserve the adjusted layout in the output file.
// Common Searches: Aspose.Cells C# load workbook with auto‑fit rows enabled | How to auto‑adjust row heights on opening an Excel file using Aspose.Cells .NET | Set LoadOptions.AutoFitterOptions to auto‑fit rows in Aspose.Cells | Preserve original layout after loading Excel with Aspose.Cells auto‑fit rows | OnlyAuto false effect on row height auto‑fit in Aspose.Cells
// Tags: auto‑fit row heights via LoadOptions Aspose.Cells | AutoFitterOptions OnlyAuto false C# example | adjust row height during workbook load .NET | preserve Excel layout after loading Aspose.Cells | load XLSX with automatic row height adjustment Aspose

using System;
using Aspose.Cells;

namespace AsposeCellsAutoFitOnLoad
{
    // The example loads an XLSX workbook with Aspose.Cells using LoadOptions that contain an AutoFitterOptions object (OnlyAuto = false), automatically adjusts all row heights on load, and saves the workbook while preserving the new layout.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string inputPath = "input.xlsx";

            // Configure load options to enable auto‑fitting of rows (and columns if desired)
            LoadOptions loadOptions = new LoadOptions();

            // Create AutoFitterOptions and set properties.
            // OnlyAuto = false ensures that all rows are auto‑fitted, even those with custom heights.
            AutoFitterOptions autoFitOptions = new AutoFitterOptions
            {
                OnlyAuto = false
                // Additional options can be set here, e.g., MaxRowHeight, IgnoreHidden, etc.
            };

            // Assign the options to the load options.
            loadOptions.AutoFitterOptions = autoFitOptions;

            // Load the workbook with the specified options.
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // (Optional) Verify that rows have been auto‑fitted.
            Console.WriteLine("Row 0 height after load: " + workbook.Worksheets[0].Cells.GetRowHeight(0));

            // Save the workbook preserving the auto‑fitted layout.
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved with auto‑fitted rows to: " + outputPath);
        }
    }
}
