// Title: Aspose.Cells for .NET – List Worksheet PageSetup: Paper Size Mode, PaperSize, FitToPagesWide & FitToPagesTall (C#)
// Description: This C# example creates a workbook, adds three worksheets, configures manual paper size and FitToPages settings on two sheets, then iterates through all worksheets to output each sheet's page‑setup mode (automatic or manual), PaperSize, FitToPagesWide, and FitToPagesTall. The report is printed to the console and the workbook is saved as an XLSX file.
// Keywords: Aspose.Cells | C# | .NET | Worksheet PageSetup | PaperSize | FitToPagesWide | FitToPagesTall | print layout report | automatic paper size | manual paper size
// Common Searches: Aspose.Cells get worksheet paper size | How to read FitToPagesWide in Aspose.Cells | C# list page setup for all worksheets | Aspose.Cells page setup report example | Print layout settings Aspose.Cells .NET
// Developer Intent: Generate a console report that enumerates each worksheet’s page‑setup configuration—including mode, paper size, FitToPagesWide, and FitToPagesTall—and save the workbook.
// Use Cases: Audit printing settings before converting the workbook to PDF | Validate that all worksheets comply with corporate print‑layout standards | Log page‑setup details for debugging workbook generation | Create a summary sheet that documents layout configurations for end users
// AI Prompts: Write a method that returns a List of objects containing worksheet name, paper mode, paper size, FitToPagesWide, and FitToPagesTall using Aspose.Cells. | Modify the sample to export the page‑setup report to a CSV file with column headers instead of writing to the console. | Add comprehensive error handling for missing or null PageSetup objects and include inline comments explaining each step. | Show how to enable IsAutomaticPaperSize, reflect the change in the report, and explain the impact on printing.

using System;
using Aspose.Cells;

namespace AsposeCellsReport
{
    // This C# example creates a workbook, adds three worksheets, configures manual paper size and FitToPages settings on two sheets, then iterates through all worksheets to output each sheet's page‑setup mode (automatic or manual), PaperSize, FitToPagesWide, and FitToPagesTall. The report is printed to the console and the workbook is saved as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (using the provided create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet and set some page setup values
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Summary";
            // Default settings are kept for sheet1

            // Add a second worksheet and customize its page setup
            int sheetIndex = workbook.Worksheets.Add();
            Worksheet sheet2 = workbook.Worksheets[sheetIndex];
            sheet2.Name = "Details";

            // Set paper size to A5 and fit to 2 pages wide, height adjusts automatically
            sheet2.PageSetup.PaperSize = PaperSizeType.PaperA5;
            sheet2.PageSetup.SetFitToPages(2, 0); // Fit to 2 pages wide, height auto

            // Add a third worksheet with different settings
            int sheetIndex3 = workbook.Worksheets.Add();
            Worksheet sheet3 = workbook.Worksheets[sheetIndex3];
            sheet3.Name = "Report";

            // Set paper size to Letter and fit to a single page both ways
            sheet3.PageSetup.PaperSize = PaperSizeType.PaperLetter;
            sheet3.PageSetup.FitToPagesWide = 1;
            sheet3.PageSetup.FitToPagesTall = 1;

            // Generate the report: list each worksheet's paper size mode, FitToPagesWide, and FitToPagesTall
            Console.WriteLine("Worksheet Page Setup Report");
            Console.WriteLine("----------------------------");
            foreach (Worksheet ws in workbook.Worksheets)
            {
                PageSetup ps = ws.PageSetup;

                // Paper size mode: automatic if IsAutomaticPaperSize is true, otherwise manual
                string paperMode = ps.IsAutomaticPaperSize ? "Automatic" : "Manual";

                // Current paper size (only meaningful when mode is Manual)
                PaperSizeType paperSize = ps.PaperSize;

                // FitToPages settings
                int fitWide = ps.FitToPagesWide;
                int fitTall = ps.FitToPagesTall;

                Console.WriteLine($"Worksheet: {ws.Name}");
                Console.WriteLine($"  Paper Size Mode : {paperMode}");
                Console.WriteLine($"  Paper Size      : {paperSize}");
                Console.WriteLine($"  FitToPagesWide  : {fitWide}");
                Console.WriteLine($"  FitToPagesTall  : {fitTall}");
                Console.WriteLine();
            }

            // Save the workbook (using the provided save rule)
            workbook.Save("PageSetupReport.xlsx", SaveFormat.Xlsx);
        }
    }
}
