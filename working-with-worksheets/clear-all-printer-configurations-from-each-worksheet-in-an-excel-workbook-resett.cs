using System;
using Aspose.Cells;

class ClearPrinterSettings
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Reset printer settings to the default (null clears any custom settings)
            sheet.PageSetup.PrinterSettings = null;

            // Optionally clear other print-related properties to ensure defaults
            sheet.PageSetup.PrintArea = null;
            sheet.PageSetup.PrintTitleRows = null;
            sheet.PageSetup.PrintTitleColumns = null;
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}