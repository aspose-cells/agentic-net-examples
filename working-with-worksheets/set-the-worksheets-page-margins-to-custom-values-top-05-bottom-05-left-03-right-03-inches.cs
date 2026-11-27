// Title: How to set custom page margins (top 0.5", bottom 0.5", left 0.3", right 0.3") for an Excel worksheet using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set worksheet top margin to 0.5 inches | C# define custom left and right margins for Excel print layout with Aspose.Cells | How to configure page margins in an Excel file using Aspose.Cells PageSetup | Set bottom margin to 0.5 inches in Aspose.Cells workbook C# | Adjust Excel worksheet margins in inches via Aspose.Cells API
// Tags: Aspose.Cells PageSetup margin configuration | C# set Excel worksheet print margins | custom inches margins Excel Aspose.Cells | worksheet page layout margins .NET | Excel workbook margin settings using Aspose

using Aspose.Cells;
using System;

// The example creates a new Workbook, accesses the first Worksheet, assigns 0.5 inches to TopMargin and BottomMargin and 0.3 inches to LeftMargin and RightMargin through the PageSetup object, and saves the workbook as CustomMargins.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set custom page margins (values are in inches)
        sheet.PageSetup.TopMargin = 0.5;      // Top margin
        sheet.PageSetup.BottomMargin = 0.5;   // Bottom margin
        sheet.PageSetup.LeftMargin = 0.3;     // Left margin
        sheet.PageSetup.RightMargin = 0.3;    // Right margin

        // Save the workbook to a file
        workbook.Save("CustomMargins.xlsx");
    }
}
