// Title: Set worksheet page margins (0.5" top/bottom, 0.3" left/right) with Aspose.Cells for .NET
// Description: Demonstrates how to assign custom top, bottom, left, and right margins in inches to a worksheet using Aspose.Cells PageSetup properties, then save the workbook as an Excel file.
// Keywords: Aspose.Cells page margins | C# set worksheet margins | TopMarginInch | BottomMarginInch | LeftMarginInch | RightMarginInch | Excel print margins inches | custom worksheet margins | Aspose.Cells .NET example
// Common Searches: Aspose.Cells set custom page margins .NET | How to change worksheet margins in inches using Aspose.Cells | TopMarginInch BottomMarginInch Aspose.Cells example | Set left and right margins for Excel sheet with Aspose.Cells | Print layout margins Aspose.Cells C#
// Developer Intent: Apply precise top, bottom, left, and right margin values to a worksheet via Aspose.Cells PageSetup before saving or exporting.
// Use Cases: Create print‑ready reports that must adhere to specific inch margins. | Prepare invoice templates matching corporate letterhead specifications. | Generate PDFs from Excel with exact margin settings for publishing. | Standardize margin layout across multiple worksheets in a workbook.
// AI Prompts: Write C# code using Aspose.Cells to set 0.5" top/bottom and 0.3" left/right margins and export the workbook to PDF. | Explain how to retrieve current margin values from a worksheet's PageSetup in Aspose.Cells. | Create a reusable method that accepts margin parameters and applies them to every worksheet in a given Aspose.Cells workbook. | Show how to convert inches to points when setting margins with Aspose.Cells for non‑inch units.

using System;
using Aspose.Cells;

namespace AsposeCellsMarginExample
{
    // Demonstrates how to assign custom top, bottom, left, and right margins in inches to a worksheet using Aspose.Cells PageSetup properties, then save the workbook as an Excel file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set custom page margins (in inches)
            worksheet.PageSetup.TopMarginInch = 0.5;     // Top margin
            worksheet.PageSetup.BottomMarginInch = 0.5;  // Bottom margin
            worksheet.PageSetup.LeftMarginInch = 0.3;    // Left margin
            worksheet.PageSetup.RightMarginInch = 0.3;   // Right margin

            // Optionally add some data to visualize the margins
            worksheet.Cells["A1"].PutValue("Margin Demo");
            worksheet.Cells["A2"].PutValue("Top: 0.5\", Bottom: 0.5\", Left: 0.3\", Right: 0.3\"");

            // Save the workbook to a file
            workbook.Save("CustomMargins.xlsx");
        }
    }
}
