using System;
using Aspose.Cells;

class RemovePrinterSettingsDemo
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample data for printer settings test");

        // (Optional) Assign dummy printer settings to demonstrate removal later
        sheet.PageSetup.PrinterSettings = new byte[] { 0x01, 0x02, 0x03 };

        // Remove printer settings by setting the property to null
        sheet.PageSetup.PrinterSettings = null;

        // Save the workbook; when opened in Excel it should not prompt for a default printer
        string outputPath = "Workbook_NoPrinterPrompt.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        // Load the saved workbook to verify that printer settings are indeed cleared
        Workbook loadedWorkbook = new Workbook(outputPath);
        bool printerSettingsExist = loadedWorkbook.Worksheets[0].PageSetup.PrinterSettings != null;
        Console.WriteLine("PrinterSettings present after reload: " + printerSettingsExist);
    }
}