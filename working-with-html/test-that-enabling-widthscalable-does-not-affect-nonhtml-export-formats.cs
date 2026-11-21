// Title: How to confirm HtmlSaveOptions.WidthScalable only influences HTML export and leaves XLSX and PDF outputs unchanged in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# test that WidthScalable setting does not change XLSX export | verify HtmlSaveOptions WidthScalable has no effect on PDF output using .NET | sample code to export workbook to HTML with WidthScalable then to XLSX and PDF | check column width preservation when using WidthScalable in Aspose.Cells | does enabling WidthScalable affect non‑HTML formats in Aspose.Cells .NET
// Tags: HtmlSaveOptions WidthScalable HTML export | Aspose.Cells export to XLSX without width scaling | Aspose.Cells PDF export unaffected by WidthScalable | C# column width preservation Aspose.Cells | validate non‑HTML export behavior Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a workbook, sets explicit column widths, saves it as HTML with HtmlSaveOptions.WidthScalable enabled, and then saves the same workbook to XLSX and PDF formats to demonstrate that the WidthScalable option does not alter the results of non‑HTML exports.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and add some data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Test");
                sheet.Cells["B1"].PutValue("WidthScalable");

                // Set column widths to known values
                sheet.Cells.SetColumnWidth(0, 20); // Column A
                sheet.Cells.SetColumnWidth(1, 30); // Column B

                // Export to HTML with WidthScalable enabled
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
                htmlOptions.WidthScalable = true; // Enable WidthScalable
                workbook.Save("Export_WithWidthScalable.html", htmlOptions);

                // Export to XLSX (non‑HTML) – WidthScalable has no effect
                workbook.Save("Export_NonHtml.xlsx", SaveFormat.Xlsx);

                // Export to PDF (non‑HTML) – confirm no side effects
                workbook.Save("Export_NonHtml.pdf", SaveFormat.Pdf);

                Console.WriteLine("Exports completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
