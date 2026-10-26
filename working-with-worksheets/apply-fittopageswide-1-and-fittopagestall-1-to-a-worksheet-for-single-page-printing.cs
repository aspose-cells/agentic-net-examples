// Title: Set FitToPagesWide and FitToPagesTall to 1 for single‑page printing of a worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to fit an entire worksheet on one printed page | set FitToPagesWide = 1 and FitToPagesTall = 1 with Aspose.Cells .NET | single page print layout for Excel using Aspose.Cells API | C# code to configure page setup for one‑page print in Aspose.Cells | Aspose.Cells page setup example for printing to a single page
// Tags: Aspose.Cells page setup fit to one page | C# set FitToPagesWide property | C# set FitToPagesTall property | single page printing Excel Aspose.Cells | worksheet page layout configuration Aspose.Cells

using Aspose.Cells;
using System;

// The example creates or loads a Workbook, accesses its first Worksheet, and configures the PageSetup by setting FitToPagesWide and FitToPagesTall to 1, forcing the worksheet to print on a single page. The workbook is then saved as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // default workbook with one worksheet

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Configure the page setup to print the entire sheet on a single page
            sheet.PageSetup.FitToPagesWide = 1;   // fit to 1 page wide
            sheet.PageSetup.FitToPagesTall = 1;   // fit to 1 page tall

            // Save the workbook
            string outputPath = "SinglePagePrint.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
