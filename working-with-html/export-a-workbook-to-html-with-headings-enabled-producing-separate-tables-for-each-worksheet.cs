// Title: Export Aspose.Cells Workbook to Separate HTML Files with Row/Column Headings (C#)
// Description: Demonstrates how to create a workbook with multiple worksheets, enable row and column headings, and save each sheet as an individual HTML file using Aspose.Cells HtmlSaveOptions in .NET.
// Keywords: Aspose.Cells HTML export | C# export workbook to HTML | row column headings Aspose | multiple HTML files per worksheet | HtmlSaveOptions SaveAsSingleFile false | ExportActiveWorksheetOnly false
// Common Searches: Aspose.Cells export workbook to HTML with headings | save each worksheet as separate HTML file C# | HtmlSaveOptions ExportRowColumnHeadings example | how to disable single file output Aspose.Cells | export multiple sheets to individual HTML pages
// Developer Intent: Generate separate HTML files for every worksheet while preserving row and column headers.
// Use Cases: Create paginated HTML reports where each sheet appears on its own page with full headers. | Provide web‑ready data tables for a multi‑page portal, keeping column and row titles intact. | Export Excel data to HTML for documentation or email distribution, with each worksheet isolated.
// AI Prompts: Show how to specify an output folder and custom file naming when exporting multiple worksheets to HTML with Aspose.Cells. | Add custom CSS to the generated HTML tables while retaining row and column headings. | Explain how to export only selected worksheets to separate HTML files using HtmlSaveOptions.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to create a workbook with multiple worksheets, enable row and column headings, and save each sheet as an individual HTML file using Aspose.Cells HtmlSaveOptions in .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add data to the first worksheet (default sheet)
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";
            sheet1.Cells["A1"].PutValue("Header1");
            sheet1.Cells["B1"].PutValue("Header2");
            sheet1.Cells["A2"].PutValue("Data1");
            sheet1.Cells["B2"].PutValue("Data2");

            // Add a second worksheet and populate it
            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
            sheet2.Cells["A1"].PutValue("ColA");
            sheet2.Cells["B1"].PutValue("ColB");
            sheet2.Cells["A2"].PutValue("Value1");
            sheet2.Cells["B2"].PutValue("Value2");

            // Configure HTML save options
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                // Enable export of row and column headings
                ExportRowColumnHeadings = true,

                // Ensure each worksheet is saved as a separate HTML file (default behavior)
                SaveAsSingleFile = false,

                // Export all worksheets (not only the active one)
                ExportActiveWorksheetOnly = false
            };

            // Save the workbook to HTML.
            // When SaveAsSingleFile is false, Aspose.Cells creates separate HTML files for each sheet.
            // The base file name is used as a prefix for each generated file.
            workbook.Save("WorkbookOutput.html", options);
        }
    }
}
