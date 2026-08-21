// Title: Aspose.Cells .NET – Freeze Top Row, Set Print Title Row, Render Pages for Verification
// Description: Creates a workbook with 100 rows, freezes the first row, assigns it as the repeatable print title, defines a print area, saves the file, checks freeze‑pane settings, and renders each printed page to PNG images to confirm the title row appears on every page.
// Keywords: Aspose.Cells freeze panes | print title rows Aspose.Cells | repeat header on printed pages .NET | render worksheet pages to images | verify print titles programmatically | Aspose.Cells PageSetup PrintArea | C# Aspose.Cells example
// Common Searches: how to freeze first row and repeat as print title in Aspose.Cells | render printed Excel pages to PNG with Aspose.Cells .NET | check frozen panes and print titles via code | set print title rows and print area together Aspose.Cells | verify print header repeats on each page
// Developer Intent: Freeze the top row, designate it as the print title that repeats on every printed page, and programmatically confirm the configuration by rendering the pages.
// Use Cases: Generate a spreadsheet, freeze the header row, and set it as a repeating print title. | Define a specific print area and save the workbook for distribution. | Retrieve and display freeze‑pane parameters and the PrintTitleRows setting for validation. | Render each printed page to an image file to visually verify that the title row repeats across pages.
// AI Prompts: Write C# code using Aspose.Cells to freeze the first row, set it as the print title row, define a print area, and render each printed page to PNG for verification. | Explain how to programmatically validate frozen panes and print title rows in an Aspose.Cells workbook and output the results to the console. | Provide tips for configuring ImageOrPrintOptions so that the header row appears on every rendered page when using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPrintTitleWithFreezeDemo
{
    // Creates a workbook with 100 rows, freezes the first row, assigns it as the repeatable print title, defines a print area, saves the file, checks freeze‑pane settings, and renders each printed page to PNG images to confirm the title row appears on every page.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data (100 rows, 5 columns)
            for (int row = 0; row < 100; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Freeze the first row so it stays visible while scrolling
            // FreezePanes(row, column, freezedRows, freezedColumns)
            // Freeze at cell A2 (row index 1) with 1 frozen row and 0 frozen columns
            worksheet.FreezePanes(2, 1, 1, 0);

            // Set the first row as the print title row (will repeat on each printed page)
            worksheet.PageSetup.PrintTitleRows = "$1:$1";

            // Define the print area to include all data
            worksheet.PageSetup.PrintArea = "A1:E100";

            // Save the workbook (lifecycle rule)
            workbook.Save("PrintTitleRowsWithFreeze.xlsx");

            // Verify that the freeze panes are set correctly
            bool hasFreeze = worksheet.GetFreezedPanes(out int freezeRow, out int freezeColumn,
                                                       out int frozenRows, out int frozenColumns);
            Console.WriteLine($"Freeze panes set: {hasFreeze}");
            if (hasFreeze)
            {
                Console.WriteLine($"Freeze position - Row: {freezeRow}, Column: {freezeColumn}");
                Console.WriteLine($"Frozen rows: {frozenRows}, Frozen columns: {frozenColumns}");
            }

            // Verify that the print title rows property is set
            Console.WriteLine($"PrintTitleRows = {worksheet.PageSetup.PrintTitleRows}");

            // Render each printed page to an image to visually confirm that the title row appears on every page
            ImageOrPrintOptions printOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = false, // allow multiple pages
                HorizontalResolution = 96,
                VerticalResolution = 96
            };

            SheetRender sheetRender = new SheetRender(worksheet, printOptions);
            int pageCount = sheetRender.PageCount;
            Console.WriteLine($"Total pages to be printed: {pageCount}");

            for (int i = 0; i < pageCount; i++)
            {
                string imagePath = $"PrintedPage_{i + 1}.png";
                sheetRender.ToImage(i, imagePath);
                Console.WriteLine($"Rendered page {i + 1} to {imagePath}");
            }

            // Clean up resources
            sheetRender.Dispose();

            Console.WriteLine("Demo completed. Verify the generated images to ensure the title row repeats on each page.");
        }
    }
}
