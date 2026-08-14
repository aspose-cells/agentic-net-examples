// Title: Automatic Page Breaks in Aspose.Cells for .NET Using GetPrintingPageBreaks
// Description: C# example that creates a workbook, fills 200 rows, defines a print area, sets FitToPagesWide = 1 and FitToPagesTall = 0, retrieves the automatically generated page breaks with GetPrintingPageBreaks, optionally inserts matching horizontal page breaks, and saves the file for preview in Excel.
// Keywords: Aspose.Cells | automatic page breaks | GetPrintingPageBreaks | FitToPagesTall zero | horizontal page breaks | C# | .NET | print area | Excel pagination | worksheet printing
// Common Searches: Aspose.Cells get automatic page breaks .NET | How to enable dynamic page breaks in Excel with Aspose.Cells | FitToPagesTall = 0 pagination example | Add horizontal page breaks programmatically Aspose.Cells | Retrieve printed page ranges C# Aspose
// Developer Intent: Programmatically obtain and apply automatic page breaks based on worksheet content height.
// Use Cases: Determine the row range for each printed page to split a workbook into separate PDFs. | Insert explicit horizontal page breaks that mirror the automatic pagination before distributing the file. | Adjust print settings to keep the sheet one page wide while allowing the height to break automatically.
// AI Prompts: Generate C# code that exports each automatically detected page to an individual PDF using Aspose.Cells. | Show how to log the start and end rows of every page break to a text file instead of the console. | Explain how to set custom row heights so that automatic page breaks occur at predetermined rows.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AutomaticPageBreakDemo
{
    // C# example that creates a workbook, fills 200 rows, defines a print area, sets FitToPagesWide = 1 and FitToPagesTall = 0, retrieves the automatically generated page breaks with GetPrintingPageBreaks, optionally inserts matching horizontal page breaks, and saves the file for preview in Excel.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate the worksheet with enough rows to require multiple pages when printed
            for (int i = 0; i < 200; i++)
            {
                // Fill column A with sample text
                worksheet.Cells[i, 0].PutValue($"Row {i + 1}");
                // Add some extra data to make rows taller (optional)
                worksheet.Cells[i, 1].PutValue($"Data {i + 1}");
            }

            // Configure page setup:
            // - Set a print area covering all populated rows
            // - Allow the height to adjust automatically by setting FitToPagesTall to 0
            // - Keep width fitting to one page for clarity
            PageSetup pageSetup = worksheet.PageSetup;
            pageSetup.PrintArea = "A1:B200";
            pageSetup.FitToPagesWide = 1;   // one page wide
            pageSetup.FitToPagesTall = 0;   // height is not forced, automatic page breaks will be used

            // Create print options (default options are sufficient for page break calculation)
            ImageOrPrintOptions options = new ImageOrPrintOptions();

            // Retrieve automatic page breaks based on the current content and page setup
            CellArea[] automaticPageBreaks = worksheet.GetPrintingPageBreaks(options);

            // Output information about each automatic page break
            Console.WriteLine($"Total automatic page breaks detected: {automaticPageBreaks.Length}");
            for (int i = 0; i < automaticPageBreaks.Length; i++)
            {
                CellArea area = automaticPageBreaks[i];
                // Each CellArea represents the range of cells that will be printed on a single page
                Console.WriteLine($"Page {i + 1}: Starts at Row {area.StartRow + 1}, Column {area.StartColumn + 1} " +
                                  $"- Ends at Row {area.EndRow + 1}, Column {area.EndColumn + 1}");
            }

            // (Optional) If you want to make the page breaks explicit in the worksheet,
            // add horizontal page breaks at the end row of each automatic page area.
            foreach (CellArea area in automaticPageBreaks)
            {
                // Add a horizontal page break after the last row of the page area
                // The Add method with a single row parameter adds a break at the top-left of that row.
                worksheet.HorizontalPageBreaks.Add(area.EndRow + 1);
            }

            // Save the workbook to verify the layout (the file can be opened in Excel to see the page breaks)
            workbook.Save("AutomaticPageBreakDemo.xlsx");

            Console.WriteLine("Workbook saved. Open 'AutomaticPageBreakDemo.xlsx' to view the automatic page breaks.");
        }
    }
}
