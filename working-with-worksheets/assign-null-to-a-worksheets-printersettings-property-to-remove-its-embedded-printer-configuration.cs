// Title: Clear worksheet printer settings in Aspose.Cells (C#) by setting PageSetup.PrinterSettings to null
// Description: Demonstrates how to remove an embedded printer configuration from a worksheet by assigning null to its PageSetup.PrinterSettings property, then saving the workbook.
// Keywords: Aspose.Cells clear printer settings | PageSetup.PrinterSettings null | remove worksheet printer configuration | reset Excel printer settings C# | Aspose.Cells workbook without printer data
// Common Searches: how to clear printer settings in Aspose.Cells | set worksheet PrinterSettings to null C# | remove embedded printer configuration from Excel file | Aspose.Cells reset page setup printer | clear printer configuration before saving workbook
// Developer Intent: Remove the worksheet's embedded printer configuration by assigning null to its PageSetup.PrinterSettings property.
// Use Cases: Prepare a template that must not retain any printer‑specific data before distribution. | Generate reports programmatically and ensure they open correctly on any printer. | Strip printer settings from a workbook to avoid conflicts when shared across different environments.
// AI Prompts: Write C# code that clears printer settings for every worksheet in an existing Aspose.Cells workbook. | Explain the impact of setting PageSetup.PrinterSettings to null on the saved Excel file and any known limitations. | Show how to programmatically verify that printer settings have been removed after the workbook is saved.

using System;
using Aspose.Cells;

namespace AsposeCellsPrinterSettingsDemo
{
    // Demonstrates how to remove an embedded printer configuration from a worksheet by assigning null to its PageSetup.PrinterSettings property, then saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Optionally add some data (not required for printer settings)
            worksheet.Cells["A1"].PutValue("Demo: Remove printer settings");

            // Access the page setup of the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // Assign null to the PrinterSettings property to clear embedded printer configuration
            pageSetup.PrinterSettings = null;

            // Verify that the property is now null (optional)
            Console.WriteLine("PrinterSettings is null: " + (pageSetup.PrinterSettings == null));

            // Save the workbook (lifecycle: save)
            workbook.Save("WorkbookWithoutPrinterSettings.xlsx");

            Console.WriteLine("Workbook saved successfully.");
        }
    }
}
