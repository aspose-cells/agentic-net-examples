// Title: Export only a defined print area from an Aspose.Cells workbook to HTML in C#
// Common Searches: C# Aspose.Cells export print area only to HTML file | How to save a specific range as HTML using Aspose.Cells HtmlSaveOptions | Aspose.Cells HtmlSaveOptions ExportPrintAreaOnly property example in .NET | Export worksheet range A1:C3 to HTML with grid lines using Aspose.Cells | Save active worksheet as HTML with only the print area in C#
// Tags: Aspose.Cells HtmlSaveOptions ExportPrintAreaOnly | export worksheet print area as HTML | save workbook to HTML with grid lines | set worksheet print area A1:C3 Aspose.Cells | export active worksheet only to HTML

using System;
using Aspose.Cells;

namespace ExportPrintAreaToHtml
{
    // The example creates a workbook, defines a print area (A1:C3), configures HtmlSaveOptions with ExportPrintAreaOnly, ExportGridLines, ExportHeadings, and ExportActiveWorksheetOnly, and saves the workbook as an HTML file that contains only the specified print area.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the output HTML file
            string outputHtmlPath = "PrintAreaExport.html";

            try
            {
                // ------------------------------------------------------------
                // 1. Create a new workbook and add some sample data
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DataSheet";

                // Populate cells with sample data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["C1"].PutValue("Price");

                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["C2"].PutValue(0.5);

                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["C3"].PutValue(0.3);

                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(15);
                sheet.Cells["C4"].PutValue(0.8);

                // ------------------------------------------------------------
                // 2. Define the print area (e.g., only the first three rows)
                // ------------------------------------------------------------
                // The print area is defined using the PageSetup.PrintArea property.
                // Here we set it to A1:C3, which includes the header row and two data rows.
                sheet.PageSetup.PrintArea = "A1:C3";

                // ------------------------------------------------------------
                // 3. Configure HtmlSaveOptions to export only the print area
                // ------------------------------------------------------------
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // ExportPrintAreaOnly = true ensures that only the defined print area
                    // is written to the resulting HTML file.
                    ExportPrintAreaOnly = true,

                    // Optional: Export grid lines and headings for better readability.
                    ExportGridLines = true,
                    ExportHeadings = true,

                    // Optional: Export the active worksheet only (since we have a single sheet).
                    ExportActiveWorksheetOnly = true
                };

                // ------------------------------------------------------------
                // 4. Save the workbook as HTML using the configured options
                // ------------------------------------------------------------
                workbook.Save(outputHtmlPath, htmlOptions);

                Console.WriteLine($"HTML file successfully saved to '{outputHtmlPath}'.");
                Console.WriteLine("Only the defined print area (A1:C3) was exported.");
            }
            catch (Exception ex)
            {
                // Output any errors that occur during processing
                Console.WriteLine("An error occurred while exporting the print area to HTML:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
