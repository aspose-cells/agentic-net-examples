// Title: Set narrow print margins on an Excel worksheet using Aspose.Cells in C#
// AI Prompts: Write C# code that uses Aspose.Cells to set the left, right, top, and bottom print margins to 0.25 inches for a worksheet. | Show how to configure the PageSetup object of an Aspose.Cells worksheet to shrink print margins for maximum printable area. | Provide a complete example that creates or loads a workbook, applies 0.25‑inch print margins, and saves the result with Aspose.Cells.
// Common Searches: Aspose.Cells C# how to change worksheet page margins to inches | reduce Excel print margins with Aspose.Cells library | C# code sample for setting narrow margins in an Excel file using Aspose | maximize printable area in Excel by adjusting margins via Aspose.Cells | set custom left and right margins for a worksheet in Aspose.Cells C#
// Tags: Aspose.Cells worksheet PageSetup margins | C# set Excel print margins inches | narrow margin configuration Aspose.Cells | Excel printable area optimization C# | adjust worksheet margins programmatically

using Aspose.Cells;

// The program creates (or loads) an Excel workbook, accesses the first worksheet, sets all four print margins to 0.25 inches via the PageSetup object, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // Replace with new Workbook("input.xlsx") to load

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set narrow print margins (values are in inches)
        sheet.PageSetup.LeftMargin = 0.25;
        sheet.PageSetup.RightMargin = 0.25;
        sheet.PageSetup.TopMargin = 0.25;
        sheet.PageSetup.BottomMargin = 0.25;

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
