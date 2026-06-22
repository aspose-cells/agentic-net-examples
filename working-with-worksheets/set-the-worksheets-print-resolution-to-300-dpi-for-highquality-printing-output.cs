using System;
using Aspose.Cells;

class SetPrintResolution
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the print resolution (print quality) to 300 DPI
        worksheet.PageSetup.PrintQuality = 300;

        // Output the current setting to verify
        Console.WriteLine("Print Quality set to: " + worksheet.PageSetup.PrintQuality + " DPI");

        // Save the workbook
        workbook.Save("PrintResolution300DPI.xlsx");
    }
}