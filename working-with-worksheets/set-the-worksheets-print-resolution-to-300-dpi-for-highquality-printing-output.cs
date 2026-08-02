// Title: Set Worksheet Print Resolution to 300 DPI with Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, accesses the first Worksheet, sets its PageSetup.PrintQuality to 300 DPI, verifies the value, and saves the file as "PrintResolution300DPI.xlsx"—ideal for high‑resolution printing and PDF/image export.
// Keywords: Aspose.Cells | C# | PrintQuality | 300 DPI | worksheet print resolution | PageSetup | high quality printing | Excel export | PDF image quality
// Common Searches: Aspose.Cells set worksheet DPI | C# set print quality 300 DPI Aspose.Cells | How to change Excel print resolution using Aspose.Cells | PageSetup PrintQuality 300 DPI example | Set high resolution print settings Aspose.Cells .NET
// Developer Intent: Configure a worksheet's print resolution to 300 DPI for high‑resolution output.
// Use Cases: Generate Excel reports that must print at 300 DPI for professional publishing. | Prepare workbooks for conversion to PDF or image formats where crisp print quality is required. | Standardize print settings across multiple worksheets before distributing to clients or printers.
// AI Prompts: Show C# code that sets PrintQuality to 300 DPI for every worksheet in an Aspose.Cells workbook. | Give an example that reads back the PrintQuality value after setting it and logs the result. | Explain how to combine PrintQuality with other PageSetup options such as margins and orientation for high‑quality prints.

using System;
using Aspose.Cells;

// Creates a new Workbook, accesses the first Worksheet, sets its PageSetup.PrintQuality to 300 DPI, verifies the value, and saves the file as "PrintResolution300DPI.xlsx"—ideal for high‑resolution printing and PDF/image export.
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

        // Verify the setting
        Console.WriteLine("Print Quality set to: " + worksheet.PageSetup.PrintQuality + " DPI");

        // Save the workbook
        workbook.Save("PrintResolution300DPI.xlsx");
    }
}
