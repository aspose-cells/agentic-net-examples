// Title: Load an Excel workbook without charts, assign an InterruptMonitor, and export to PDF using Aspose.Cells for .NET
// AI Prompts: Load an .xlsx file with Aspose.Cells LoadOptions that skips chart objects, set a workbook.InterruptMonitor, and save the workbook as a PDF. | Use reflection to apply the LoadDataOnly property on LoadOptions only when it exists, then convert the workbook to PDF while ignoring charts.
// Common Searches: Aspose.Cells load workbook without loading charts and convert to PDF in C# | How to set InterruptMonitor on a workbook in Aspose.Cells .NET | Conditional use of LoadDataOnly property with Aspose.Cells LoadOptions | Export Excel to PDF while ignoring chart objects using Aspose.Cells | C# example for loading Excel data only and saving as PDF with Aspose.Cells
// Tags: skip chart objects LoadOptions Aspose.Cells | configure workbook InterruptMonitor Aspose.Cells | reflection set LoadDataOnly property C# | export Excel workbook to PDF Aspose.Cells | load data only mode Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks for the input Excel file, creates a LoadOptions instance, uses reflection to enable LoadDataOnly when available, loads the workbook while skipping chart objects, assigns an InterruptMonitor, and saves the workbook as a PDF, handling any errors that may occur.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputFile = "input.xlsx";
                const string outputFile = "output.pdf";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Create LoadOptions; enable LoadDataOnly if the property exists in the current version
                LoadOptions loadOptions = new LoadOptions();
                var loadDataOnlyProp = typeof(LoadOptions).GetProperty("LoadDataOnly");
                if (loadDataOnlyProp != null && loadDataOnlyProp.CanWrite)
                {
                    loadDataOnlyProp.SetValue(loadOptions, true);
                }

                // Load the workbook with the specified options
                Workbook workbook = new Workbook(inputFile, loadOptions);

                // Assign an interrupt monitor (optional; no interruption is triggered here)
                workbook.InterruptMonitor = new InterruptMonitor();

                // Save the workbook as PDF
                workbook.Save(outputFile, SaveFormat.Pdf);
                Console.WriteLine($"Workbook successfully saved as PDF: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
