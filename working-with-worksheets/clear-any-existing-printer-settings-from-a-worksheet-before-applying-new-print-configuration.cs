// Title: Clear worksheet printer settings before applying new print configuration with Aspose.Cells for .NET
// Description: Demonstrates how to reset a worksheet's printer settings by assigning null to PageSetup.PrinterSettings, then configure draft mode and a specific print area, and finally save the workbook using Aspose.Cells in C#.
// Keywords: Aspose.Cells C# clear printer settings | PageSetup.PrinterSettings null | reset worksheet printer configuration | Aspose.Cells print draft mode | set print area Aspose.Cells | Excel workbook printer reset | Aspose.Cells .NET printing example | clear printer settings before printing | programmatic Excel print setup | Aspose.Cells PageSetup example
// Common Searches: How to reset printer settings on an Excel worksheet using Aspose.Cells | Aspose.Cells clear PageSetup.PrinterSettings C# | Remove previous printer configuration before setting print area in .NET | Set draft mode and print area after clearing printer settings with Aspose.Cells | Clear worksheet printer settings programmatically
// Developer Intent: Remove any existing printer configuration from a worksheet so that new print options can be applied without inheriting old settings.
// Use Cases: Prepare a template workbook for multiple reports that require different print layouts. | Ensure consistent draft‑mode printing for generated invoices. | Avoid conflicts when reusing a worksheet that previously had custom printer settings.
// AI Prompts: Generate C# code with Aspose.Cells that clears a worksheet's printer settings and then sets PrintDraft and a custom PrintArea. | Explain the effect of assigning null to PageSetup.PrinterSettings in Aspose.Cells and why it is needed before redefining print options. | Show how to verify that printer settings have been cleared before saving the workbook using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to reset a worksheet's printer settings by assigning null to PageSetup.PrinterSettings, then configure draft mode and a specific print area, and finally save the workbook using Aspose.Cells in C#.
    public class ClearPrinterSettingsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the PageSetup object for the worksheet
                PageSetup pageSetup = worksheet.PageSetup;

                // Clear any existing printer settings
                pageSetup.PrinterSettings = null;

                // Apply new print configuration as needed
                pageSetup.PrintDraft = true;               // Print without graphics (draft mode)
                pageSetup.PrintArea = "A1:C10";            // Define the range to be printed

                // Save the workbook to verify the changes
                string outputPath = "ClearedPrinterSettings.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ClearPrinterSettingsDemo.Run();
        }
    }
}
