// Title: Set narrow (0.25‑inch) print margins on an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells to set the left, right, top, and bottom page margins of a worksheet to 0.25 inches and then save the file. | Demonstrate how to configure PageSetup margins to the minimum printable area in an Excel workbook using Aspose.Cells in a .NET application. | Create an Excel workbook where all print margins are set to 0.25 inches via Aspose.Cells and export it as NarrowMargins.xlsx.
// Common Searches: Aspose.Cells C# set worksheet print margins to 0.25 inches | how to reduce Excel print margins programmatically with Aspose.Cells .NET | maximizing printable area in Excel by using narrow margins Aspose.Cells | C# example for PageSetup margin adjustment using Aspose.Cells | custom print margins for Excel workbook Aspose.Cells API
// Tags: Aspose.Cells set worksheet page margins | C# PageSetup margin configuration | narrow print margins Excel workbook | maximize printable area Aspose.Cells | custom left right top bottom margins .NET | save workbook with custom margins Aspose.Cells

using Aspose.Cells;

// The program creates a new workbook, accesses the first worksheet, sets all page margins to 0.25 inches to maximize printable area, and saves the file as NarrowMargins.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set narrow print margins (in inches) to maximize printable area
        sheet.PageSetup.LeftMargin = 0.25;
        sheet.PageSetup.RightMargin = 0.25;
        sheet.PageSetup.TopMargin = 0.25;
        sheet.PageSetup.BottomMargin = 0.25;

        // Save the workbook
        workbook.Save("NarrowMargins.xlsx");
    }
}
