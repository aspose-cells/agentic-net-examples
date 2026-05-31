using System;
using Aspose.Cells;

namespace AsposeCellsPrinterSettingsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the page setup of the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // Remove any embedded printer configuration by assigning null
            pageSetup.PrinterSettings = null;

            // Verify that the printer settings are cleared
            Console.WriteLine("PrinterSettings cleared: " + (pageSetup.PrinterSettings == null));

            // Save the workbook (lifecycle rule: save)
            workbook.Save("PrinterSettingsRemoved.xlsx");

            Console.WriteLine("Workbook saved successfully.");
        }
    }
}