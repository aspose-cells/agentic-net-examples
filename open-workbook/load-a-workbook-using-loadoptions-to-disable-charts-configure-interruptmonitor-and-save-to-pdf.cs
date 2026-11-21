// Title: Load an XLSX workbook without charts using LoadOptions and InterruptMonitor, then export to PDF in C# with Aspose.Cells
// Common Searches: C# Aspose.Cells load XLSX without charts using LoadOptions | How to set a timeout for workbook conversion with InterruptMonitor in Aspose.Cells | Skip chart data when opening an Excel file with LoadDataOnly in Aspose.Cells | Convert Excel to PDF in C# while disabling chart rendering with Aspose.Cells | Use LoadOptions and InterruptMonitor together for fast PDF export in Aspose.Cells
// Tags: LoadOptions chart suppression | InterruptMonitor timeout configuration | Workbook to PDF conversion Aspose.Cells | disable charts on workbook load | C# Excel to PDF export Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example checks for the input XLSX file, creates LoadOptions for the XLSX format, enables LoadDataOnly to skip loading charts, loads the workbook, configures an InterruptMonitor with a 30‑second timeout (if supported), assigns it to the workbook, and finally saves the workbook as a PDF, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load options for XLSX format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Set LoadDataOnly if the property exists in the current version
            var loadDataOnlyProp = typeof(LoadOptions).GetProperty("LoadDataOnly");
            if (loadDataOnlyProp != null && loadDataOnlyProp.CanWrite)
            {
                loadDataOnlyProp.SetValue(loadOptions, true);
            }

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Configure an InterruptMonitor (set timeout if supported)
            InterruptMonitor monitor = new InterruptMonitor();
            var timeoutProp = typeof(InterruptMonitor).GetProperty("Timeout");
            if (timeoutProp != null && timeoutProp.CanWrite)
            {
                timeoutProp.SetValue(monitor, 30000); // timeout in milliseconds
            }
            workbook.InterruptMonitor = monitor;

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
