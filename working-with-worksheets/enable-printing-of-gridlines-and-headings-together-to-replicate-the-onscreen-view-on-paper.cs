// Title: Print gridlines and row/column headings together in an Excel worksheet with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# enable printing of gridlines and headings on the same page | How to set PrintGridlines and PrintHeadings in Aspose.Cells .NET | Fit worksheet to one page width while printing gridlines using Aspose.Cells | C# code to print Excel row and column headings with gridlines via Aspose.Cells | Aspose.Cells page setup print gridlines and headings together
// Tags: PageSetup gridlines printing | PageSetup headings printing | FitToPagesWide printing option | Aspose.Cells save workbook Xlsx | Worksheet print configuration C#

using System;
using Aspose.Cells;

// The example creates a workbook, accesses the first worksheet, enables printing of both gridlines and row/column headings through PageSetup properties, optionally fits the content to a single page width, and saves the result as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Enable printing of gridlines
        sheet.PageSetup.PrintGridlines = true;

        // Enable printing of row and column headings (A, B, 1, 2, …)
        sheet.PageSetup.PrintHeadings = true;

        // Optional: fit the content to one page width while allowing height to expand
        sheet.PageSetup.FitToPagesWide = 1;
        sheet.PageSetup.FitToPagesTall = 0;

        // Save the workbook
        workbook.Save("PrintGridlinesAndHeadings.xlsx", SaveFormat.Xlsx);
    }
}
