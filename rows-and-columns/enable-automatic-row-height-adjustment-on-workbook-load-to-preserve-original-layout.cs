// Title: Auto‑fit all row heights on workbook load with Aspose.Cells for .NET
// Description: Shows how to configure LoadOptions and AutoFitterOptions (OnlyAuto = false) so that every row is auto‑fitted when a workbook is opened, then saved with the adjusted heights.
// Keywords: Aspose.Cells | auto fit rows | LoadOptions | AutoFitterOptions | OnlyAuto | C# | row height adjustment | preserve layout | Excel load | auto‑fit on load
// Common Searches: Aspose.Cells auto fit rows on load | How to auto‑adjust row height when opening an Excel file in C# | LoadOptions AutoFitterOptions OnlyAuto example | Auto‑fit rows without losing custom heights Aspose.Cells | C# code to auto‑fit rows during workbook load
// Developer Intent: Resize all rows automatically during workbook loading to retain the original visual layout.
// Use Cases: Open a template workbook and ensure rows match their content before adding data. | Process a batch of reports, preserving row heights so the final PDFs look identical to the source files. | Load user‑provided spreadsheets, auto‑fit rows, and then export them to another format without manual adjustments.
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells and auto‑fits every row using LoadOptions.AutoFitterOptions. | Explain the effect of the OnlyAuto property on row height adjustment when a workbook is loaded. | Show how to combine AutoFitterOptions with other LoadOptions settings (e.g., PreserveFormulas) in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AutoFitOnLoadDemo
{
    // Shows how to configure LoadOptions and AutoFitterOptions (OnlyAuto = false) so that every row is auto‑fitted when a workbook is opened, then saved with the adjusted heights.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string inputPath = "input.xlsx";

            // Configure load options to auto‑fit rows when the workbook is loaded
            LoadOptions loadOptions = new LoadOptions();

            // Create AutoFitterOptions and set properties as needed.
            // OnlyAuto = false ensures that all rows are auto‑fitted,
            // not only those without a custom height.
            loadOptions.AutoFitterOptions = new AutoFitterOptions
            {
                OnlyAuto = false
            };

            // Load the workbook with the specified options.
            // The rows will be auto‑fitted during this operation.
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // (Optional) Verify that a row height has been adjusted.
            double firstRowHeight = workbook.Worksheets[0].Cells.GetRowHeight(0);
            Console.WriteLine($"First row height after load auto‑fit: {firstRowHeight}");

            // Save the workbook to preserve the adjusted row heights.
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved with auto‑fitted rows to: {outputPath}");
        }
    }
}
