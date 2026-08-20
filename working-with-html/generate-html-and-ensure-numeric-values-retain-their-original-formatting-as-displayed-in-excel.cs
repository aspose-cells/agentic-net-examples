// Title: C# – Export Excel to HTML5 while preserving numeric, date & percentage formats with Aspose.Cells
// Description: Demonstrates how to create a workbook, apply custom number, date and percentage formats, configure HtmlSaveOptions to keep column widths and styling, save the sheet as HTML5, and retrieve a single cell's HTML string using Aspose.Cells for .NET.
// Keywords: Aspose.Cells HTML export C# | preserve Excel number format HTML | custom date format Aspose.Cells | percentage formatting HTML5 | keep column width Excel to HTML | GetHtmlString Aspose.Cells | Excel to HTML5 conversion .NET | Aspose.Cells HtmlSaveOptions
// Common Searches: Aspose.Cells keep thousand separator when converting to HTML | export Excel date cell with custom format to HTML C# | Aspose.Cells preserve percentage formatting in HTML output | HTML5 export from Excel workbook using Aspose.Cells .NET | how to retain column widths in Aspose.Cells HTML export
// Developer Intent: Generate an HTML5 file from an Excel workbook that displays numbers, dates and percentages exactly as they appear in Excel.
// Use Cases: Financial web reports that require thousand separators and two‑decimal precision (e.g., 12,345.68). | Public dashboards showing dates in a specific pattern such as 15‑Aug‑2023. | Performance metrics displayed as percentages (e.g., 25.67%) without losing Excel styling. | Embedding a single formatted cell into a web page via GetHtmlString.
// AI Prompts: Write C# code with Aspose.Cells to export a worksheet to HTML5, keeping custom number, date, and percentage formats and original column widths. | Show how to retrieve the HTML5 markup of a single formatted cell using Aspose.Cells GetHtmlString. | Explain which HtmlSaveOptions properties are essential for preserving Excel styling when saving to HTML with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExportDemo
{
    // Demonstrates how to create a workbook, apply custom number, date and percentage formats, configure HtmlSaveOptions to keep column widths and styling, save the sheet as HTML5, and retrieve a single cell's HTML string using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook and get the first worksheet
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // 2. Populate cells with numeric values and apply custom formats
            // ------------------------------------------------------------
            // Simple number with thousand separator and two decimals
            cells["A1"].PutValue(12345.6789);
            Style styleA1 = cells["A1"].GetStyle();
            styleA1.Custom = "#,##0.00";               // e.g., 12,345.68
            cells["A1"].SetStyle(styleA1);

            // Date value with a specific display format
            cells["A2"].PutValue(new DateTime(2023, 8, 15));
            Style styleA2 = cells["A2"].GetStyle();
            styleA2.Custom = "dd-mmm-yyyy";            // e.g., 15-Aug-2023
            cells["A2"].SetStyle(styleA2);

            // Percentage value
            cells["A3"].PutValue(0.2567);
            Style styleA3 = cells["A3"].GetStyle();
            styleA3.Custom = "0.00%";                  // e.g., 25.67%
            cells["A3"].SetStyle(styleA3);

            // ------------------------------------------------------------
            // 3. Configure HTML save options to preserve formatting
            // ------------------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Use HTML5 for better compatibility (optional)
                HtmlVersion = HtmlVersion.Html5,

                // Export all data (including styles) – default is All
                ExportDataOptions = HtmlExportDataOptions.All,

                // Do NOT ignore column width; keep Excel‑like truncation behavior
                FormatDataIgnoreColumnWidth = false,

                // Export the displayed value, not the formula text
                ExportFormula = false,

                // Keep numeric formatting as shown in Excel
                // (no extra property needed – the style applied to cells is respected)
            };

            // ------------------------------------------------------------
            // 4. Save the workbook as an HTML file
            // ------------------------------------------------------------
            string htmlPath = "FormattedNumbers.html";
            workbook.Save(htmlPath, htmlOptions);
            Console.WriteLine($"HTML file saved to: {htmlPath}");

            // ------------------------------------------------------------
            // 5. (Optional) Retrieve the HTML string of a single cell
            // ------------------------------------------------------------
            string cellHtml = cells["A1"].GetHtmlString(true); // true => HTML5 wrapper
            Console.WriteLine("HTML representation of A1:");
            Console.WriteLine(cellHtml);
        }
    }
}
