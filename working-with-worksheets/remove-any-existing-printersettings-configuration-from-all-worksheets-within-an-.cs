using System;
using Aspose.Cells;

class RemovePrinterSettingsDemo
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Iterate through all worksheets and clear the printer settings
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // The PrinterSettings property is a byte array; setting it to null removes any existing configuration
            sheet.PageSetup.PrinterSettings = null;
        }

        // Save the workbook after removing printer settings
        string outputPath = "output_without_printer_settings.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}