// Title: C# – Export Excel to HTML while preserving custom numeric formats with Aspose.Cells
// Description: Demonstrates how to create a workbook, apply custom number formats (thousand separator, percentage, scientific notation), configure HtmlSaveOptions to keep Excel styling and column widths, save as HTML, and retrieve a cell's formatted HTML string using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | HTML export | custom number format | preserve Excel formatting | HtmlSaveOptions | GetHtmlString | thousand separator | percentage format | scientific notation | column width retention | HTML5
// Common Searches: Aspose.Cells export Excel to HTML keep number format | C# preserve thousand separator in HTML export | percentage format HTML output Aspose.Cells | scientific notation HTML conversion Excel | GetHtmlString formatted cell Aspose.Cells | HtmlSaveOptions column width Excel preview
// Developer Intent: Generate an HTML file from an Excel workbook that retains the exact numeric display and layout defined in Excel.
// Use Cases: Create web‑ready reports that show numbers exactly as they appear in Excel (e.g., commas, % signs, scientific notation). | Provide an on‑line preview of an Excel sheet where column widths and cell styles match the original file. | Extract a single cell’s HTML markup for embedding in emails, dashboards, or other web components.
// AI Prompts: Show C# code using Aspose.Cells to export a workbook to HTML while preserving custom numeric formats and column widths. | Explain how to set HtmlSaveOptions so that thousand separators, percentages, and scientific notation are retained in the HTML output. | Demonstrate retrieving the HTML representation of a formatted cell with GetHtmlString for web integration.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to create a workbook, apply custom number formats (thousand separator, percentage, scientific notation), configure HtmlSaveOptions to keep Excel styling and column widths, save as HTML, and retrieve a cell's formatted HTML string using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate cells with numeric values
            cells["A1"].PutValue(12345.6789);
            cells["A2"].PutValue(0.256);
            cells["A3"].PutValue(987654321);

            // Apply custom number formats to retain original Excel display
            Style style = workbook.CreateStyle();

            // A1: thousand separator with two decimals
            style.Custom = "#,##0.00";
            cells["A1"].SetStyle(style);

            // A2: percentage with one decimal
            style = workbook.CreateStyle();
            style.Custom = "0.0%";
            cells["A2"].SetStyle(style);

            // A3: scientific notation
            style = workbook.CreateStyle();
            style.Custom = "0.00E+00";
            cells["A3"].SetStyle(style);

            // Configure HTML save options to preserve formatting
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Use HTML5 for better compatibility
                HtmlVersion = HtmlVersion.Html5,

                // Export all data (including styles) to HTML
                ExportDataOptions = HtmlExportDataOptions.All,

                // Ensure column width truncation behaves like Excel
                FormatDataIgnoreColumnWidth = false
            };

            // Save the workbook as an HTML file
            workbook.Save("FormattedNumbers.html", htmlOptions);

            // Optional: demonstrate retrieving the HTML string of a cell
            string cellHtml = cells["A1"].GetHtmlString(true);
            Console.WriteLine("HTML representation of A1:");
            Console.WriteLine(cellHtml);
        }
    }
}
