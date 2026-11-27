// Title: Set narrow print margins on an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set worksheet print margins to 0.25 inches | How to reduce page margins in Excel file with Aspose.Cells .NET | C# code to maximize printable area by adjusting margins in Aspose.Cells | Configure left, right, top, bottom margins in Aspose.Cells PageSetup | Example of setting narrow margins for Excel workbook using Aspose.Cells
// Tags: Aspose.Cells worksheet print margin configuration | C# PageSetup margin adjustment inches | Excel narrow margins Aspose.Cells .NET | maximize printable area Excel Aspose.Cells | set left right top bottom margins Aspose.Cells

using Aspose.Cells;

// Creates a new workbook, accesses the first worksheet, sets all page margins to 0.25 inches to maximize printable area, and saves the file as NarrowMargins.xlsx.
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
