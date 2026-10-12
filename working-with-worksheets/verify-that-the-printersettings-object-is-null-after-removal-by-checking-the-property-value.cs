// Title: Remove printer settings from a worksheet and confirm they are null with Aspose.Cells for .NET
// AI Prompts: Write C# code that sets Worksheet.PageSetup.PrinterSettings to null and returns a boolean indicating the result. | Generate a console application using Aspose.Cells that clears the printer configuration of the first worksheet and prints whether the property is null. | Create a method that removes printer settings from a workbook's worksheet and validates the removal by checking the PrinterSettings property.
// Common Searches: aspnet how to clear printer settings on an Excel worksheet using Aspose.Cells | c# verify that Worksheet.PageSetup.PrinterSettings becomes null after assignment | example code to remove printer configuration from Aspose.Cells workbook and check null value
// Tags: clear worksheet printer settings Aspose.Cells | validate printer settings null Aspose.Cells | Aspose.Cells PageSetup PrinterSettings removal | C# remove printer configuration Excel workbook

using System;
using Aspose.Cells;

// // Demonstrates creating a workbook, setting the first worksheet's PageSetup.PrinterSettings to null, and confirming the property is null by outputting a boolean result.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet's page setup (contains printer settings)
            Worksheet sheet = workbook.Worksheets[0];
            var initialSettings = sheet.PageSetup.PrinterSettings; // capture current settings

            // "Remove" the printer settings by setting the property to null
            sheet.PageSetup.PrinterSettings = null;

            // Verify that the printer settings are now null
            bool isPrinterSettingsNull = sheet.PageSetup.PrinterSettings == null;

            Console.WriteLine("PrinterSettings is null after removal: " + isPrinterSettingsNull);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
