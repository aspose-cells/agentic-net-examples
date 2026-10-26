// Title: Configure Aspose.Cells worksheet to fit printed output to 1 page wide and 2 pages tall (single‑column layout) in C#
// Common Searches: Aspose.Cells C# set FitToPagesWide = 1 and FitToPagesTall = 2 | how to create a single‑column print layout for a worksheet using Aspose.Cells | page scaling example with Aspose.Cells workbook in .NET | set worksheet page setup to fit to specific number of pages Aspose.Cells | print Excel sheet as 1 page wide by 2 pages tall with Aspose.Cells API
// Tags: worksheet page setup fit to pages Aspose.Cells | FitToPagesWide property C# | FitToPagesTall property Aspose.Cells | worksheet printing page count control | Aspose.Cells print scaling options

using Aspose.Cells;

// The sample creates a new Workbook, accesses its first Worksheet, configures the PageSetup to fit the printed output to 1 page wide and 2 pages tall (creating a single‑column layout), and saves the file as FitToPages.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set single‑column layout: 1 page wide, 2 pages tall
        sheet.PageSetup.FitToPagesWide = 1;   // Fit to 1 page wide
        sheet.PageSetup.FitToPagesTall = 2;   // Fit to 2 pages tall

        // Save the workbook
        workbook.Save("FitToPages.xlsx");
    }
}
