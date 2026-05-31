using System;
using Aspose.Cells;

namespace AsposeCellsPrinterSettingsDemo
{
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

            // Verify that printer settings are set
            Console.WriteLine("PrinterSettings set? " + (pageSetup.PrinterSettings != null));

            // Remove printer settings by assigning null
            pageSetup.PrinterSettings = null;

            // Verify that printer settings are now null
            bool isNull = pageSetup.PrinterSettings == null;
            Console.WriteLine("PrinterSettings after removal is null? " + isNull);

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            workbook.Save("PrinterSettingsDemo.xlsx");
        }
    }
}