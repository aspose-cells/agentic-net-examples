// Title: How to set FitToPagesTall = 1 to fit all rows on a single printed page with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set worksheet FitToPagesTall to 1 for single-page print | Fit all rows on one printed page using Aspose.Cells PageSetup | C# code to scale Excel worksheet vertically to one page with Aspose.Cells | Configure page setup to fit rows to one page in Aspose.Cells .NET | Print Excel sheet with all rows on one page using Aspose.Cells API
// Tags: FitToPagesTall property Aspose.Cells | worksheet vertical page fit Aspose.Cells | single-page row printing .NET | Aspose.Cells page setup scaling | C# Excel print scaling to one page

using Aspose.Cells;

// The example loads an existing workbook, accesses the first worksheet, sets its PageSetup.FitToPagesTall property to 1 so that all rows are compressed onto a single printed page, and then saves the workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (or any specific worksheet you need)
        Worksheet sheet = workbook.Worksheets[0];

        // Configure the page setup to fit all rows on a single printed page
        // Setting FitToPagesTall to 1 forces the printer to scale the sheet so that
        // all rows are contained within one page vertically.
        sheet.PageSetup.FitToPagesTall = 1;

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
