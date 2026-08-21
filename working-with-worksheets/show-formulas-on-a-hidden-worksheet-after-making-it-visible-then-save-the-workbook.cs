// Title: Aspose.Cells for .NET – Show Formulas on a Hidden Worksheet and Save the Workbook
// Description: C# example that creates a workbook, adds a hidden sheet, writes a SUM formula to A1, unhides the sheet, enables the ShowFormulas view, and saves the file as XLSX using Aspose.Cells.
// Keywords: Aspose.Cells C# ShowFormulas | display formulas hidden worksheet | unhide worksheet Aspose.Cells | save workbook with formulas visible | Aspose.Cells HideWorksheet property | Excel formula view programmatically | Aspose.Cells .NET example
// Common Searches: how to show formulas on a hidden sheet with Aspose.Cells | Aspose.Cells make hidden worksheet visible C# | set ShowFormulas property before saving workbook | C# Aspose.Cells display formulas instead of values | unhide worksheet and show formulas Aspose.Cells .NET
// Developer Intent: Unhide a worksheet, display its formulas, and persist the workbook.
// Use Cases: Audit hidden calculation sheets by revealing formulas before distribution. | Debug template helpers: hide formula sheets during normal use, then expose them for troubleshooting. | Automate a quality‑check step that toggles visibility and switches to formula view to verify correctness prior to final save.
// AI Prompts: Generate C# code with Aspose.Cells that unhides a worksheet, sets ShowFormulas = true, and saves the workbook as XLSX. | Provide an Aspose.Cells example that hides a sheet, inserts a formula, then makes the sheet visible and shows formulas before saving. | Explain the effect of the ShowFormulas property on the saved Excel file when a hidden worksheet is made visible.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, adds a hidden sheet, writes a SUM formula to A1, unhides the sheet, enables the ShowFormulas view, and saves the file as XLSX using Aspose.Cells.
    public class ShowFormulasOnHiddenSheetDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a new worksheet that will be hidden initially
                Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");

                // Hide the worksheet
                hiddenSheet.IsVisible = false;

                // Put a formula in cell A1 of the hidden worksheet
                hiddenSheet.Cells["A1"].Formula = "=SUM(10,20,30)";

                // Make the hidden worksheet visible
                hiddenSheet.IsVisible = true;

                // Enable showing formulas instead of their calculated results
                hiddenSheet.ShowFormulas = true;

                // Define output file path
                string outputPath = "ShowFormulasOnHiddenSheet.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}' with formulas shown on the previously hidden sheet.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ShowFormulasOnHiddenSheetDemo.Run();
        }
    }
}
