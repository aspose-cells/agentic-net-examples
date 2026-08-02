// Title: Set distinct odd and even page headers with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to enable separate headers for odd and even pages in an Excel workbook using PageSetup.IsHFDiffOddEven, SetHeader, and SetEvenHeader, then saves the file as OddEvenHeaders.xlsx.
// Keywords: Aspose.Cells odd even headers | C# PageSetup IsHFDiffOddEven | SetEvenHeader Aspose.Cells | different headers per page Excel | Aspose.Cells header customization .NET
// Common Searches: Aspose.Cells set odd page header C# | How to use SetEvenHeader in Aspose.Cells | Enable different headers for odd and even pages .NET | PageSetup IsHFDiffOddEven example
// Developer Intent: Create an Excel workbook where odd‑page and even‑page headers have separate text.
// Use Cases: Printed reports that need a title on odd pages and a logo on even pages. | Multi‑page invoices with distinct header layouts for front and back sides. | Booklet‑style workbooks where facing pages display mirrored headers for easier navigation.
// AI Prompts: Show a C# example that adds different footers for odd and even pages with Aspose.Cells. | Explain how to toggle IsHFDiffOddEven and assign custom header and footer strings in a worksheet. | Provide code to style odd‑page and even‑page headers with different fonts or colors using Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to enable separate headers for odd and even pages in an Excel workbook using PageSetup.IsHFDiffOddEven, SetHeader, and SetEvenHeader, then saves the file as OddEvenHeaders.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        PageSetup pageSetup = worksheet.PageSetup;

        // Enable different headers for odd and even pages
        pageSetup.IsHFDiffOddEven = true;

        // Set header for odd pages (regular header)
        pageSetup.SetHeader(0, "&LOdd Page Left");
        pageSetup.SetHeader(1, "&COdd Page Center");
        pageSetup.SetHeader(2, "&ROdd Page Right");

        // Set header for even pages
        pageSetup.SetEvenHeader(0, "&LEven Page Left");
        pageSetup.SetEvenHeader(1, "&CEven Page Center");
        pageSetup.SetEvenHeader(2, "&REven Page Right");

        // Save the workbook
        workbook.Save("OddEvenHeaders.xlsx");
    }
}
