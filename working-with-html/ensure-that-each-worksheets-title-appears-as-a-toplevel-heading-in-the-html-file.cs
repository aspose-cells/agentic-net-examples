using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add a few worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Name = "Summary";
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Cells["A1"].PutValue("Summary data");

            Worksheet sheet2 = workbook.Worksheets.Add("Details");
            sheet2.Cells["A1"].PutValue("Detail data");

            Worksheet sheet3 = workbook.Worksheets.Add("Statistics");
            sheet3.Cells["A1"].PutValue("Statistics data");

            // Set each worksheet's header to display its name (sheet title)
            // The header will be rendered as a top‑level heading in the HTML output
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Center section (index 1) shows the sheet name using the &A placeholder
                ws.PageSetup.SetHeader(1, "&A");
            }

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export the page header (which now contains the sheet name)
                ExportPageHeaders = true,

                // Save all sheets into a single HTML file so each header appears as a separate heading
                SaveAsSingleFile = true,
                ShowAllSheets = true
            };

            // Save the workbook as HTML
            string outputPath = "WorkbookWithSheetHeadings.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to: {outputPath}");
        }
    }
}