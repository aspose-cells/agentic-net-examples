// Title: C# – Load HTML into Aspose.Cells, add comments, and re‑export to HTML with tooltip (title) attributes
// Description: Shows how to import an HTML file into an Aspose.Cells Workbook, add or modify cell comments, set HtmlSaveOptions (IsExportComments = true, ExportCommentsType = PrintInPlace) and save the workbook back to HTML so each comment appears as a mouse‑over tooltip (title attribute).
// Keywords: Aspose.Cells | C# | HTML to Excel conversion | export comments as tooltip | HtmlSaveOptions | IsExportComments | PrintCommentsType.PrintInPlace | cell comments tooltip | load HTML workbook | re‑export HTML with tooltips
// Common Searches: Aspose.Cells export comments as HTML tooltip | C# convert HTML to Excel keep comments | HtmlSaveOptions IsExportComments true example | PrintInPlace comment export Aspose.Cells | load HTML file into Aspose.Cells workbook
// Developer Intent: Create an Excel workbook from an HTML document and save it back to HTML while preserving cell comments as title‑attribute tooltips.
// Use Cases: Import an existing HTML report, add business notes as cell comments, and generate a new HTML file where each note shows as a hover tooltip. | Extract tooltip data from HTML attributes, store them as worksheet comments for further analysis, then re‑export the workbook to HTML with the original tooltip behavior intact. | Batch‑process a collection of HTML files into Excel workbooks, ensuring that any comment information is retained as HTML title attributes when the files are saved again as HTML.
// AI Prompts: Write C# code using Aspose.Cells to load an HTML file, add a comment to cell A1, and save the workbook as HTML with comments exported as title‑attribute tooltips. | Explain the role of HtmlSaveOptions.IsExportComments and ExportCommentsType = PrintCommentsType.PrintInPlace in preserving cell comments as HTML tooltips. | Provide a step‑by‑step guide for converting HTML to Excel, adding comments, and re‑exporting to HTML while keeping tooltip functionality.

using System;
using Aspose.Cells;

// Shows how to import an HTML file into an Aspose.Cells Workbook, add or modify cell comments, set HtmlSaveOptions (IsExportComments = true, ExportCommentsType = PrintInPlace) and save the workbook back to HTML so each comment appears as a mouse‑over tooltip (title attribute).
class Program
{
    static void Main()
    {
        // Load the source HTML file into a workbook.
        // Aspose.Cells can directly load HTML documents.
        Workbook workbook = new Workbook("input.html");

        // Access the first worksheet (or any specific worksheet as needed).
        Worksheet worksheet = workbook.Worksheets[0];

        // OPTIONAL: Demonstrate adding a comment that will be exported as a tooltip.
        // In a real scenario, comments could be extracted from HTML attributes and added here.
        int commentIdx = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIdx];
        comment.Note = "This is a tooltip comment";

        // Configure HTML save options to export comments as tooltip attributes.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.IsExportComments = true;                         // Enable comment export.
        htmlOptions.ExportCommentsType = PrintCommentsType.PrintInPlace; // Export as tooltip (title attribute).

        // Save the workbook back to HTML, preserving comments as tooltips.
        workbook.Save("output.html", htmlOptions);
    }
}
