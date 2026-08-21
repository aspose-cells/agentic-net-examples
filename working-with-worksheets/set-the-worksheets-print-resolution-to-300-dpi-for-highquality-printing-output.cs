// Title: Set Worksheet Print Resolution to 300 DPI with Aspose.Cells for .NET (C#)
// Description: C# example that sets the PageSetup.PrintQuality property of a worksheet to 300 DPI, guaranteeing high‑quality printed output, and saves the workbook as an Excel file.
// Keywords: Aspose.Cells | C# print resolution | Worksheet DPI | PageSetup.PrintQuality | 300 DPI Excel | high quality printing | Aspose.Cells .NET | set print quality | Excel print DPI | Aspose.Cells example
// Common Searches: Aspose.Cells set worksheet DPI | C# set print quality 300 DPI Aspose.Cells | PageSetup.PrintQuality property example | How to change Excel print resolution using Aspose.Cells | PrintResolutionDemo Aspose.Cells
// Developer Intent: Configure a worksheet’s print resolution to 300 DPI for high‑quality printing using Aspose.Cells in C#.
// Use Cases: Create printable reports that require professional 300 DPI output. | Standardize print quality across all worksheets in a generated Excel workbook. | Develop a template that automatically applies 300 DPI when users print the sheet.
// AI Prompts: Generate C# code with Aspose.Cells that sets a worksheet’s print resolution to 300 DPI and saves the file. | Explain the effect of the PageSetup.PrintQuality property on printed Excel output and how to read its value programmatically. | Show how to apply a 300 DPI print resolution to every worksheet in an existing workbook using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsPrintResolutionDemo
{
    // C# example that sets the PageSetup.PrintQuality property of a worksheet to 300 DPI, guaranteeing high‑quality printed output, and saves the workbook as an Excel file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the print resolution (print quality) to 300 DPI
            worksheet.PageSetup.PrintQuality = 300;

            // Verify the setting (optional)
            Console.WriteLine("Print Quality set to: " + worksheet.PageSetup.PrintQuality + " DPI");

            // Save the workbook
            workbook.Save("PrintResolutionDemo.xlsx");
        }
    }
}
