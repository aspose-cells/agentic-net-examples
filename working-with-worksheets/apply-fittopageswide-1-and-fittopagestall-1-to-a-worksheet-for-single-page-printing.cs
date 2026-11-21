// Title: Set worksheet to print on a single page using FitToPagesWide and FitToPagesTall in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set worksheet to fit on one printed page | How to use FitToPagesWide and FitToPagesTall properties in Aspose.Cells | Print Excel sheet as single page with Aspose.Cells .NET API | C# code to configure page scaling to one page in Aspose.Cells workbook | Aspose.Cells page setup example for single-page printing
// Tags: Aspose.Cells page setup scaling | FitToPagesWide Aspose.Cells C# | FitToPagesTall Aspose.Cells C# | Excel worksheet one-page print layout Aspose.Cells | C# Aspose.Cells print layout configuration

using System;
using Aspose.Cells;

// The example creates a workbook, accesses the first worksheet, sets PageSetup.FitToPagesWide and FitToPagesTall to 1 to force the sheet to print on a single page, and saves the file as SinglePagePrint.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet (index 0)
        Worksheet sheet = workbook.Worksheets[0];

        // Set the page setup to fit the entire sheet on a single page
        sheet.PageSetup.FitToPagesWide = 1;   // Fit to 1 page wide
        sheet.PageSetup.FitToPagesTall = 1;   // Fit to 1 page tall

        // Save the workbook to a file
        workbook.Save("SinglePagePrint.xlsx");
    }
}
