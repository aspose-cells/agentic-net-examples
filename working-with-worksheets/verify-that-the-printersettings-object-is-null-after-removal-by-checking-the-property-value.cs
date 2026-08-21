// Title: Verify PageSetup.PrinterSettings becomes null after clearing in Aspose.Cells for .NET
// Description: This C# example creates a workbook, accesses the first worksheet's PageSetup, assigns a dummy byte array to the PrinterSettings property, confirms the assignment, clears the property by setting it to null, checks that it is now null, and saves the file.
// Keywords: Aspose.Cells | .NET | C# | PageSetup.PrinterSettings | null check | clear printer settings | worksheet example | sample code | GitHub | API usage
// Common Searches: Aspose.Cells check if PrinterSettings is null | How to clear PageSetup.PrinterSettings in .NET | Verify removal of printer settings Aspose.Cells | C# example for resetting worksheet printer settings | Aspose.Cells null PrinterSettings after assignment
// Developer Intent: Confirm that assigning null to PageSetup.PrinterSettings removes the printer configuration and the property returns null.
// Use Cases: Ensure no printer configuration is persisted when exporting a workbook. | Create unit tests that validate printer settings are cleared before saving. | Conditionally apply default printer settings only when the property is null.
// AI Prompts: Write an NUnit test in C# that asserts PageSetup.PrinterSettings is null after setting it to null using Aspose.Cells. | Provide code to clear PrinterSettings on every worksheet in a workbook and verify the property is null for each sheet. | Explain how to programmatically reset printer settings in an Aspose.Cells workbook and confirm they are not saved.

using System;
using Aspose.Cells;

namespace AsposeCellsPrinterSettingsCheck
{
    // This C# example creates a workbook, accesses the first worksheet's PageSetup, assigns a dummy byte array to the PrinterSettings property, confirms the assignment, clears the property by setting it to null, checks that it is now null, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the PageSetup object
            PageSetup pageSetup = worksheet.PageSetup;

            // Set some dummy printer settings (byte array)
            byte[] dummySettings = new byte[5];
            pageSetup.PrinterSettings = dummySettings;

            // Verify that the printer settings are set (should be non‑null)
            Console.WriteLine("PrinterSettings set? " + (pageSetup.PrinterSettings != null));

            // Remove the printer settings by assigning null
            pageSetup.PrinterSettings = null;

            // Verify that the printer settings are now null
            bool isNull = pageSetup.PrinterSettings == null;
            Console.WriteLine("PrinterSettings after removal is null? " + isNull);

            // Save the workbook (save rule)
            workbook.Save("PrinterSettingsCheck.xlsx");
        }
    }
}
