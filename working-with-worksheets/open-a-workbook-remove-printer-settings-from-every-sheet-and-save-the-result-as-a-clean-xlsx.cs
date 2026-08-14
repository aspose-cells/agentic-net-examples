// Title: C# – Remove All Worksheet Printer Settings with Aspose.Cells and Save a Clean XLSX
// Description: Load an existing workbook using Aspose.Cells for .NET, loop through each worksheet, clear its stored printer configuration by assigning PageSetup.PrinterSettings = null, and write the result to a new XLSX file without any printer settings.
// Keywords: Aspose.Cells | C# | .NET Excel API | remove printer settings | clear PageSetup.PrinterSettings | clean Excel workbook | strip printer configuration | save as XLSX | Excel file sanitization
// Common Searches: How to clear printer settings in an Excel file using Aspose.Cells C# | Aspose.Cells remove worksheet printer configuration | Set PageSetup.PrinterSettings to null C# | Save Excel without printer settings .NET | Programmatically strip printer settings from a workbook
// Developer Intent: Open an existing Excel file, delete all printer settings from every worksheet, and save the modified workbook as a new clean XLSX document.
// Use Cases: Prepare a workbook for distribution by removing printer configurations that may cause printing errors on other machines. | Create a template that contains no printer settings before sharing it with clients or partners. | Automate batch cleaning of printer settings across multiple workbooks in a CI/CD pipeline.
// AI Prompts: Write C# code with Aspose.Cells that clears printer settings on all worksheets and saves a new XLSX file. | Explain why assigning null to PageSetup.PrinterSettings removes stored printer data in Aspose.Cells. | Show an example of iterating through a workbook's worksheets to reset printer settings while keeping other page‑setup options unchanged.

using System;
using Aspose.Cells;

// Load an existing workbook using Aspose.Cells for .NET, loop through each worksheet, clear its stored printer configuration by assigning PageSetup.PrinterSettings = null, and write the result to a new XLSX file without any printer settings.
class RemovePrinterSettings
{
    static void Main()
    {
        // Load the existing workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Remove printer settings from each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // The PrinterSettings property holds printer configuration as a byte array.
            // Setting it to null clears any stored printer settings.
            sheet.PageSetup.PrinterSettings = null;
        }

        // Save the cleaned workbook as a new XLSX file
        string outputPath = "clean.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
