// Title: Convert HTML to Excel and Preserve Cell Comments as Tooltip Attributes with Aspose.Cells for .NET
// Description: Loads an HTML file into an Aspose.Cells Workbook, saves it as XLSX, then re‑exports the workbook to HTML while exporting cell comments as tooltip attributes using HtmlSaveOptions (IsExportComments, ExportCommentsType=PrintInPlace, AddTooltipText).
// Keywords: Aspose.Cells HTML to Excel conversion | export cell comments as tooltip | HtmlSaveOptions IsExportComments | ExportCommentsType PrintInPlace | AddTooltipText Aspose.Cells | C# convert HTML to XLSX | preserve comments round‑trip | .NET Excel HTML tooltip
// Common Searches: Aspose.Cells load HTML and save as XLSX | how to export cell comments as HTML tooltips | HtmlSaveOptions tooltip comments C# | preserve comments when converting HTML to Excel | Aspose.Cells round‑trip HTML Excel with comments
// Developer Intent: Load an HTML document into a Workbook, save it as an Excel file, and export it back to HTML with each cell comment rendered as a tooltip attribute.
// Use Cases: Generate Excel reports from HTML templates while keeping annotation tooltips for web viewers. | Automate conversion of annotated HTML tables to XLSX and back, retaining comment metadata for documentation review. | Create a bidirectional conversion pipeline where comments are stored as HTML title attributes for easy hover display.
// AI Prompts: Write C# code using Aspose.Cells to load an HTML file, save it as XLSX, and export back to HTML with comments shown as tooltips. | Explain the role of HtmlSaveOptions.IsExportComments, ExportCommentsType, and AddTooltipText in preserving cell comments as HTML tooltip attributes. | Provide a step‑by‑step tutorial for converting HTML to Excel and back while keeping cell comments as tooltip attributes with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace HtmlToExcelWithComments
{
    // Loads an HTML file into an Aspose.Cells Workbook, saves it as XLSX, then re‑exports the workbook to HTML while exporting cell comments as tooltip attributes using HtmlSaveOptions (IsExportComments, ExportCommentsType=PrintInPlace, AddTooltipText).
    class Program
    {
        static void Main()
        {
            // Paths for source HTML and destination files
            string htmlFile = "sample.html";
            string excelFile = "sample.xlsx";
            string htmlWithComments = "sample_with_comments.html";

            // -----------------------------------------------------------------
            // 1. Load the HTML file into a Workbook.
            //    LoadOptions with LoadFormat.Html tells Aspose.Cells to treat the
            //    source as an HTML document.
            // -----------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(htmlFile, loadOptions);

            // -----------------------------------------------------------------
            // 2. Save the workbook as an Excel file (XLSX).
            // -----------------------------------------------------------------
            workbook.Save(excelFile, SaveFormat.Xlsx);

            // -----------------------------------------------------------------
            // 3. (Optional) Export the same workbook back to HTML while
            //    preserving cell comments as tooltip attributes.
            //    - Enable comment export.
            //    - Use PrintInPlace so comments appear as tool‑tips.
            //    - Enable AddTooltipText to ensure tool‑tips are added.
            // -----------------------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.IsExportComments = true;                     // export comments
            htmlOptions.ExportCommentsType = PrintCommentsType.PrintInPlace; // tooltip style
            htmlOptions.AddTooltipText = true;                      // add tooltip text

            workbook.Save(htmlWithComments, htmlOptions);
        }
    }
}
