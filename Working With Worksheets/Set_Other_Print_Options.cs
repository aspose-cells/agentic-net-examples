using System;
using Aspose.Cells;

class SetOtherPrintOptionsDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data
        sheet.Cells["A1"].PutValue("Print Options Demo");

        // Get the PageSetup object to configure print settings
        PageSetup pageSetup = sheet.PageSetup;

        // Print without graphics (draft mode)
        pageSetup.PrintDraft = true;

        // Set the print quality (dots per inch)
        pageSetup.PrintQuality = 600;

        // Assign printer settings as a byte array (example placeholder)
        byte[] printerSettings = new byte[10]; // dummy data for illustration
        pageSetup.PrinterSettings = printerSettings;

        // Additional common print options
        pageSetup.BlackAndWhite = true;          // Print in black and white
        pageSetup.CenterHorizontally = true;     // Center the sheet horizontally
        pageSetup.CenterVertically = true;       // Center the sheet vertically
        pageSetup.PrintGridlines = true;         // Print cell gridlines
        pageSetup.PrintHeadings = false;         // Do not print row/column headings

        // Save the workbook with the configured print options
        workbook.Save("OtherPrintOptionsDemo.xlsx");
    }
}