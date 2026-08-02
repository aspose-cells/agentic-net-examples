// Title: Automatic Page Breaks by Content Height with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to let Aspose.Cells calculate horizontal page breaks based on worksheet content height. The sample creates a workbook, fills column A with 200 rows, defines a print area, sets FitToPagesWide = 1 and FitToPagesTall = 0, retrieves the generated breaks via GetPrintingPageBreaks, logs their row numbers, adds them to the HorizontalPageBreaks collection, and saves the file for preview in Excel.
// Keywords: Aspose.Cells | C# | automatic page breaks | GetPrintingPageBreaks | fit to page width | FitToPagesTall | horizontal page breaks programmatically | print area setup | Excel pagination | PDF export preparation
// Common Searches: Aspose.Cells get automatic page breaks C# | fit worksheet width one page height auto Aspose.Cells | add horizontal page breaks from GetPrintingPageBreaks | determine row numbers of page breaks Aspose.Cells | print area and page setup Aspose.Cells .NET
// Developer Intent: Retrieve the page breaks that Aspose.Cells computes for printing and optionally apply them to the worksheet for accurate pagination.
// Use Cases: Generate printable Excel files where rows are automatically split across pages without manual break definitions. | Synchronize programmatically added page breaks with Aspose.Cells' printing layout before converting to PDF or image formats. | Validate pagination by counting and locating automatic page breaks to ensure correct page flow in reports.
// AI Prompts: Show how to limit the maximum page height instead of using FitToPagesTall = 0. | Provide code that writes each automatic page break row number to a log file and clears existing manual breaks before adding new ones. | Explain how to customize ImageOrPrintOptions (DPI, paper size, orientation) when calling GetPrintingPageBreaks.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to let Aspose.Cells calculate horizontal page breaks based on worksheet content height. The sample creates a workbook, fills column A with 200 rows, defines a print area, sets FitToPagesWide = 1 and FitToPagesTall = 0, retrieves the generated breaks via GetPrintingPageBreaks, logs their row numbers, adds them to the HorizontalPageBreaks collection, and saves the file for preview in Excel.
class AutomaticPageBreakDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate the worksheet with enough rows to require multiple pages when printed
        for (int i = 0; i < 200; i++)
        {
            worksheet.Cells[i, 0].PutValue($"Row {i + 1}");
        }

        // Define the print area to include all populated rows
        worksheet.PageSetup.PrintArea = "A1:A200";

        // Allow the height to adjust automatically (FitToPagesTall = 0)
        // and fit the width to a single page
        worksheet.PageSetup.FitToPagesTall = 0;
        worksheet.PageSetup.FitToPagesWide = 1;

        // Create print options required by GetPrintingPageBreaks
        ImageOrPrintOptions options = new ImageOrPrintOptions();

        // Retrieve the automatically calculated page breaks based on the content height
        CellArea[] automaticPageBreaks = worksheet.GetPrintingPageBreaks(options);

        // Output information about each automatic page break
        Console.WriteLine($"Automatic page breaks count: {automaticPageBreaks.Length}");
        for (int i = 0; i < automaticPageBreaks.Length; i++)
        {
            // EndRow is zero‑based; add 1 for human‑readable row number
            Console.WriteLine($"Break {i}: Ends at row {automaticPageBreaks[i].EndRow + 1}");
        }

        // Optionally, add these automatic breaks to the worksheet's HorizontalPageBreaks collection
        // so they become visible in Excel's page break preview.
        foreach (CellArea area in automaticPageBreaks)
        {
            // Add a horizontal page break after the last row of each area
            worksheet.HorizontalPageBreaks.Add(area.EndRow);
        }

        // Save the workbook to verify the page breaks in Excel
        workbook.Save("AutomaticPageBreaks.xlsx");
    }
}
