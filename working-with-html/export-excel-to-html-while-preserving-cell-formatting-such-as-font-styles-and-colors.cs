// Title: Export Excel to HTML with full cell formatting using Aspose.Cells for .NET (C#)
// Description: Load an Excel workbook with Aspose.Cells, configure HtmlSaveOptions to keep all styles, formulas, comments and colors, then save as HTML so the output matches the original spreadsheet design.
// Keywords: Aspose.Cells HTML export | C# Excel to HTML | preserve cell styles | keep fonts colors borders | HtmlSaveOptions ExportDataOptions.All | .NET spreadsheet to HTML | export formulas to HTML
// Common Searches: Aspose.Cells export Excel to HTML with formatting | C# keep cell colors when saving workbook as HTML | HtmlSaveOptions preserve styles and formulas | How to generate HTML5 from Excel using Aspose.Cells | Export Excel workbook to styled HTML .NET
// Developer Intent: Create an HTML version of an Excel file that retains every visual element—fonts, colors, borders, and formulas—using Aspose.Cells in C#.
// Use Cases: Render a financial dashboard in a web page exactly as it appears in Excel. | Provide a read‑only HTML preview of user‑uploaded spreadsheets in a .NET web portal. | Convert Excel‑based email templates to styled HTML emails without losing design. | Generate printable HTML reports from automated Excel exports.
// AI Prompts: Generate C# code with Aspose.Cells that exports an Excel workbook to HTML while preserving fonts, colors, borders, and formulas. | Explain the impact of HtmlSaveOptions properties ExportDataOptions, ExcludeUnusedStyles, and ExportFormula on the resulting HTML. | Show a step‑by‑step guide to convert a programmatically created workbook to an HTML5 document with full styling using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace ExportExcelToHtml
{
    // Load an Excel workbook with Aspose.Cells, configure HtmlSaveOptions to keep all styles, formulas, comments and colors, then save as HTML so the output matches the original spreadsheet design.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing Excel workbook (replace with your file path)
            string excelPath = "input.xlsx";
            Workbook workbook = new Workbook(excelPath); // create/load rule

            // OPTIONAL: If you need to create a workbook from scratch, uncomment below
            // Workbook workbook = new Workbook(); // create rule
            // Worksheet sheet = workbook.Worksheets[0];
            // sheet.Cells["A1"].PutValue("Hello World");
            // Style style = sheet.Cells["A1"].GetStyle();
            // style.Font.Name = "Arial";
            // style.Font.Size = 12;
            // style.Font.Color = System.Drawing.Color.Blue;
            // sheet.Cells["A1"].SetStyle(style);

            // Configure HTML save options to preserve formatting
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // Export all data (including formulas, comments, etc.)
            htmlOptions.ExportDataOptions = HtmlExportDataOptions.All;
            // Keep all style definitions (do not exclude unused styles)
            htmlOptions.ExcludeUnusedStyles = false;
            // Ensure formulas are exported (so they appear as values in HTML)
            htmlOptions.ExportFormula = true;
            // Preserve cell coordinates if needed (optional)
            // htmlOptions.ExportCellCoordinate = true;
            // Use HTML5 standard (optional)
            // htmlOptions.HtmlVersion = HtmlVersion.Html5;

            // Save the workbook as an HTML file while preserving cell formatting
            string htmlPath = "output.html";
            workbook.Save(htmlPath, htmlOptions); // save rule

            Console.WriteLine($"Workbook exported to HTML successfully: {htmlPath}");
        }
    }
}
