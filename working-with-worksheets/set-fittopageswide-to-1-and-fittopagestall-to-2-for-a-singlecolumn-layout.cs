// Title: Set worksheet print scaling to fit 1 page wide and 2 pages tall using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set worksheet to fit 1 page wide and 2 pages tall | How to configure page setup scaling in Aspose.Cells .NET | C# example for FitToPagesWide = 1 and FitToPagesTall = 2 with Aspose.Cells | Print layout single column across page using Aspose.Cells workbook | Aspose.Cells page setup fit to pages code sample
// Tags: Aspose.Cells page setup scaling | FitToPagesWide property C# | FitToPagesTall property .NET | worksheet print layout fit to pages | save workbook as xlsx Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates or loads a workbook, accesses the first worksheet, sets its PageSetup to fit the printed output to one page wide and two pages tall, and saves the file as an XLSX document.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the PageSetup object for the worksheet
            PageSetup pageSetup = sheet.PageSetup;

            // Set the print layout to fit to 1 page wide and 2 pages tall
            pageSetup.FitToPagesWide = 1;   // one column across the page
            pageSetup.FitToPagesTall = 2;   // two rows down the page

            // Save the workbook to a file
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
