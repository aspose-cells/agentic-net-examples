// Title: C# – Verify Worksheet PageSetup.PrinterSettings Is Null After Clearing with Aspose.Cells
// Description: This example creates a Workbook, accesses the first Worksheet's PageSetup, assigns a dummy byte array to the PrinterSettings property, prints a confirmation, sets the property to null, checks that it is null, outputs the result, and saves the file. It demonstrates how to validate that printer settings have been removed using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | Worksheet PageSetup | PrinterSettings null | clear printer settings | reset page setup | unit test printer settings
// Common Searches: Aspose.Cells how to reset printer settings | PageSetup.PrinterSettings null check C# | remove worksheet printer configuration Aspose | verify printer settings cleared Aspose.Cells
// Developer Intent: Ensure that the PageSetup.PrinterSettings property returns null after it has been cleared.
// Use Cases: Unit‑test that printer settings are removed before saving a workbook. | Programmatically reset a worksheet's printer configuration to default. | Prevent unwanted printer data from being embedded in exported Excel files.
// AI Prompts: Generate C# code using Aspose.Cells that clears a worksheet's printer settings and asserts the property is null. | Explain step‑by‑step how to verify PageSetup.PrinterSettings is null after assigning null in a .NET application. | Provide a sample that removes printer settings from a worksheet, logs the verification result, and saves the workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsPrinterSettingsDemo
{
    // This example creates a Workbook, accesses the first Worksheet's PageSetup, assigns a dummy byte array to the PrinterSettings property, prints a confirmation, sets the property to null, checks that it is null, outputs the result, and saves the file. It demonstrates how to validate that printer settings have been removed using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the PageSetup object
            PageSetup pageSetup = worksheet.PageSetup;

            // Set printer settings (example byte array)
            byte[] initialSettings = new byte[5]; // dummy data
            pageSetup.PrinterSettings = initialSettings;

            // Verify that printer settings are initially set
            Console.WriteLine("PrinterSettings initially set: " + (pageSetup.PrinterSettings != null));

            // Remove printer settings by assigning null
            pageSetup.PrinterSettings = null;

            // Verify that PrinterSettings is now null
            bool isNull = pageSetup.PrinterSettings == null;
            Console.WriteLine("PrinterSettings after removal is null: " + isNull);

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            workbook.Save("PrinterSettingsDemo.xlsx");
        }
    }
}
