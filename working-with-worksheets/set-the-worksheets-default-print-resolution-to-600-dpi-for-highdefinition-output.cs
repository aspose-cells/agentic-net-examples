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

        // Set the default print resolution (print quality) to 600 DPI
        worksheet.PageSetup.PrintQuality = 600;

        // Output the set value for verification
        Console.WriteLine("Print Quality set to: " + worksheet.PageSetup.PrintQuality + " DPI");

        // Save the workbook
        workbook.Save("PrintResolution600DPI.xlsx");
    }
}