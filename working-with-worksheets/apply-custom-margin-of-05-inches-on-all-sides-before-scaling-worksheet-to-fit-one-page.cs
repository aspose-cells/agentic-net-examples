// Title: Apply 0.5‑inch page margins and fit an Excel worksheet to one page with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set page margins to 0.5 inches | How to fit an Excel worksheet on a single printed page using Aspose.Cells | C# example for custom margins before scaling worksheet in Aspose.Cells | Aspose.Cells page setup fit to one page with custom margins | Set left, right, top, bottom margins in Aspose.Cells .NET
// Tags: set page margins Aspose.Cells | fit worksheet to one page Aspose.Cells | page setup scaling Aspose.Cells .NET | custom margins Excel Aspose.Cells | print layout configuration Aspose.Cells

using Aspose.Cells;

// The sample creates a workbook, accesses the first worksheet, sets all page margins to 0.5 inches via the PageSetup object, configures the sheet to fit on a single page (both width and height), and saves the result as an Excel file.
class Program
{
    static void Main()
    {
        // {CreateWorkbook}
        Workbook workbook = new Workbook(); // placeholder for the creation rule

        // {LoadWorkbook}
        // (If loading an existing file, replace the above line with the load rule)

        Worksheet sheet = workbook.Worksheets[0];
        PageSetup setup = sheet.PageSetup;

        // Apply custom margins of 0.5 inches on all sides
        setup.LeftMargin = 0.5;
        setup.RightMargin = 0.5;
        setup.TopMargin = 0.5;
        setup.BottomMargin = 0.5;

        // Scale the worksheet to fit one page
        setup.FitToPagesWide = 1;
        setup.FitToPagesTall = 1;

        // {SaveWorkbook}
        workbook.Save("output.xlsx"); // placeholder for the save rule
    }
}
