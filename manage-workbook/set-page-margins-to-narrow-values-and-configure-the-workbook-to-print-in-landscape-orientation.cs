// Title: Set narrow page margins and print a worksheet in landscape orientation using Aspose.Cells for C#
// Common Searches: Aspose.Cells C# set worksheet margins to 0.25 inches | How to print an Excel worksheet in landscape mode with Aspose.Cells | Configure page setup margins and orientation in Aspose.Cells .NET | C# Aspose.Cells PageSetup narrow margins example | Set custom page margins and landscape orientation for Excel file using Aspose.Cells
// Tags: worksheet margin configuration Aspose.Cells C# | landscape orientation PageSetup Aspose.Cells | Aspose.Cells set margins inches | print layout settings workbook Aspose.Cells | custom page setup Aspose.Cells .NET

using Aspose.Cells;
using System;

// Creates a new workbook, accesses the first worksheet, sets all page margins to 0.25 inches, changes the orientation to landscape, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Get the page setup object for the worksheet
        PageSetup pageSetup = sheet.PageSetup;

        // Set narrow margins (values are in inches)
        pageSetup.LeftMargin = 0.25;    // left margin
        pageSetup.RightMargin = 0.25;   // right margin
        pageSetup.TopMargin = 0.25;     // top margin
        pageSetup.BottomMargin = 0.25;  // bottom margin
        pageSetup.HeaderMargin = 0.25;  // header margin
        pageSetup.FooterMargin = 0.25;  // footer margin

        // Configure the worksheet to print in landscape orientation
        pageSetup.Orientation = PageOrientationType.Landscape;

        // Save the workbook to a file
        workbook.Save("Output.xlsx", SaveFormat.Xlsx);
    }
}
