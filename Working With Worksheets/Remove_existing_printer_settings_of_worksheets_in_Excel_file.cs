using System;
using Aspose.Cells;

class RemovePrinterSettings
{
    static void Main()
    {
        // Load the existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Remove printer settings by clearing the byte array
            sheet.PageSetup.PrinterSettings = null;

            // Optionally clear any header/footer that might be tied to printer settings
            sheet.PageSetup.ClearHeaderFooter();
        }

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}