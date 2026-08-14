// Title: C# – Set custom worksheet page margins (0.5" top/bottom, 0.3" left/right) using Aspose.Cells for .NET
// Description: Creates a new Workbook, accesses the first Worksheet, uses the PageSetup object to assign TopMarginInch, BottomMarginInch, LeftMarginInch, and RightMarginInch to the required inch values, and saves the file as CustomMargins.xlsx.
// Keywords: Aspose.Cells PageSetup margins | TopMarginInch C# | BottomMarginInch example | LeftMarginInch usage | RightMarginInch code | set Excel worksheet margins programmatically | C# Excel page layout Aspose | custom worksheet margins .NET
// Common Searches: Aspose.Cells set worksheet margins in inches | C# PageSetup TopMarginInch example | How to change Excel page margins with Aspose.Cells | custom page margins Aspose.Cells for .NET | set left and right margins programmatically in Excel
// Developer Intent: Apply precise inch‑based page margin settings to a worksheet through the Aspose.Cells API.
// Use Cases: Produce printable reports that maximize data density with narrow margins. | Prepare spreadsheets for legal‑size paper where non‑standard margins are required. | Standardize margin settings before converting a workbook to PDF or XPS. | Automate workbook formatting for batch‑generated invoices with consistent layout.
// AI Prompts: Generate C# code that sets all four page margins in inches using Aspose.Cells PageSetup. | Show how to read existing worksheet margins and increase each by 0.1 inch with Aspose.Cells. | Explain how to apply identical custom margins to every sheet in an existing workbook. | Provide a snippet that saves the workbook after setting margins and then exports it to PDF.

using System;
using Aspose.Cells;

// Creates a new Workbook, accesses the first Worksheet, uses the PageSetup object to assign TopMarginInch, BottomMarginInch, LeftMarginInch, and RightMarginInch to the required inch values, and saves the file as CustomMargins.xlsx.
class SetCustomPageMargins
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the PageSetup object of the worksheet
        PageSetup pageSetup = worksheet.PageSetup;

        // Set custom margins in inches
        pageSetup.TopMarginInch = 0.5;     // Top margin = 0.5 inches
        pageSetup.BottomMarginInch = 0.5;  // Bottom margin = 0.5 inches
        pageSetup.LeftMarginInch = 0.3;    // Left margin = 0.3 inches
        pageSetup.RightMarginInch = 0.3;   // Right margin = 0.3 inches

        // Save the workbook to a file
        workbook.Save("CustomMargins.xlsx");
    }
}
