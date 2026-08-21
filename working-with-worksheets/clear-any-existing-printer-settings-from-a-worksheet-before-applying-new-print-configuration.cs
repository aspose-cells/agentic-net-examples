// Title: Reset worksheet printer settings and set new print options with Aspose.Cells for .NET (C#)
// Description: Shows how to clear existing printer settings from a worksheet by assigning null to PageSetup.PrinterSettings, then configure PrintDraft, PrintCopies, and PrintArea before saving the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells clear printer settings | PageSetup.PrinterSettings null | reset worksheet print configuration .NET | Aspose.Cells C# print draft | set print copies Aspose.Cells | define print area Aspose.Cells | reset PageSetup before printing | Aspose.Cells workbook save
// Common Searches: Aspose.Cells how to remove printer settings from a worksheet | clear PageSetup.PrinterSettings Aspose.Cells C# | reset worksheet print options before saving | set print draft and copies with Aspose.Cells | define print area in Aspose.Cells .NET | Aspose.Cells clear printer configuration for batch printing
// Developer Intent: The developer needs to delete any pre‑existing printer settings on a worksheet so that fresh print options can be applied without interference.
// Use Cases: Reuse a template workbook without inheriting prior print configurations. | Prepare workbooks for automated batch printing with consistent draft mode and copy count. | Programmatically adjust the printable range for dynamically generated reports. | Guarantee uniform printing behavior across different printers and environments.
// AI Prompts: Write C# code using Aspose.Cells to clear a worksheet's printer settings and then set PrintDraft, PrintCopies, and PrintArea. | Explain why assigning null to PageSetup.PrinterSettings resets printer configuration and list other PageSetup properties that can be modified afterward. | Create a reusable C# method that accepts draft mode, copy number, and print area parameters, clears existing printer settings, and applies the new values with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to clear existing printer settings from a worksheet by assigning null to PageSetup.PrinterSettings, then configure PrintDraft, PrintCopies, and PrintArea before saving the workbook using Aspose.Cells for .NET.
    class ClearPrinterSettingsDemo
    {
        static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the PageSetup object for the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // Clear any existing printer settings.
            pageSetup.PrinterSettings = null;

            // Apply new print configuration as needed.
            pageSetup.PrintDraft = true;          // print without graphics
            pageSetup.PrintCopies = 2;            // print two copies
            pageSetup.PrintArea = "A1:D20";       // define the area to print

            // Save the workbook
            string outputPath = "ClearedPrinterSettings.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine("Printer settings cleared and new configuration applied successfully.");
        }
    }
}
