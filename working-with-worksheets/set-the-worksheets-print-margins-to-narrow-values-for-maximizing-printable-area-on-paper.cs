// Title: Set Narrow Print Margins for an Excel Worksheet with Aspose.Cells (C#)
// Description: Creates a new Workbook, retrieves the first Worksheet, and uses the PageSetup object to assign 0.5 cm (0.2 in) margins on all sides before saving the file as NarrowMargins.xlsx.
// Keywords: Aspose.Cells C# set worksheet margins | PageSetup LeftMargin property | Excel print margins narrow | margin settings centimeters | margin settings inches | maximize printable area | Aspose.Cells .NET print setup | reduce page margins Excel | tight margins Aspose.Cells | worksheet print area optimization
// Common Searches: Aspose.Cells set worksheet margins C# | How to reduce Excel print margins with Aspose.Cells | PageSetup margin properties example | Set margins in centimeters using Aspose.Cells | Maximum printable area Excel Aspose.Cells | Configure both cm and inch margins Aspose.Cells | Narrow margins for Excel report Aspose.Cells
// Developer Intent: Apply narrow print margins to a worksheet to increase the printable area.
// Use Cases: Produce multi‑page reports that fit more rows per printed page. | Design compact invoices with minimal white space around data. | Prepare spreadsheets for small‑format paper by tightening all margins.
// AI Prompts: Show how to set 0.5 cm (0.2 in) margins for a worksheet using Aspose.Cells PageSetup in C# and save the workbook. | Provide a C# example that sets both centimeter and inch margin values for an Excel worksheet with Aspose.Cells. | Explain the difference between LeftMargin and LeftMarginInch properties and how to use them to maximize printable area.

using System;
using Aspose.Cells;

// Creates a new Workbook, retrieves the first Worksheet, and uses the PageSetup object to assign 0.5 cm (0.2 in) margins on all sides before saving the file as NarrowMargins.xlsx.
class SetNarrowMargins
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the PageSetup object for the worksheet
        PageSetup pageSetup = worksheet.PageSetup;

        // Set narrow margins (values in centimeters)
        pageSetup.LeftMargin = 0.5;    // 0.5 cm left margin
        pageSetup.RightMargin = 0.5;   // 0.5 cm right margin
        pageSetup.TopMargin = 0.5;     // 0.5 cm top margin
        pageSetup.BottomMargin = 0.5;  // 0.5 cm bottom margin

        // Also set the same margins in inches for completeness
        pageSetup.LeftMarginInch = 0.2;    // ~0.5 cm
        pageSetup.RightMarginInch = 0.2;
        pageSetup.TopMarginInch = 0.2;
        pageSetup.BottomMarginInch = 0.2;

        // Save the workbook (lifecycle save)
        workbook.Save("NarrowMargins.xlsx");
    }
}
