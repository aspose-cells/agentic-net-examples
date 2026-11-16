// Title: Export an entire Excel workbook to HTML in C# with Aspose.Cells while keeping original worksheet order and default IFilePathProvider settings
// Common Searches: c# aspnet export all worksheets to a single html file using aspose.cells preserving sheet order | how to save excel workbook as html with full path links and without changing IFilePathProvider in Aspose.Cells | Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly false example c# | generate html from excel preserving worksheet sequence Aspose.Cells C# | save workbook to html using default IFilePathProvider mapping Aspose.Cells
// Tags: export workbook to html aspnet | htmlsaveoptions exportactiveworksheetonly false | preserve worksheet order aspose.cells | default ifilepathprovider behavior | full path links html output

using System;
using Aspose.Cells;

// Loads an Excel file with Aspose.Cells, configures HtmlSaveOptions to export the full workbook (ExportActiveWorksheetOnly = false) and to generate full‑path links (IsFullPathLink = true) while leaving the IFilePathProvider mapping unchanged, then saves the result as an HTML file preserving the original worksheet order.
class ExportWorkbookToHtml
{
    static void Main()
    {
        // Path to the source Excel file.
        string excelPath = "input.xlsx";

        // Path where the HTML file will be saved.
        string htmlPath = "output.html";

        // Load the workbook without modifying any internal IFilePathProvider mappings.
        Workbook workbook = new Workbook(excelPath);

        // Configure HTML save options.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Export the entire workbook (all worksheets) preserving the original order.
            ExportActiveWorksheetOnly = false,

            // Keep the default IFilePathProvider behavior for external resources.
            // No changes to the mapping are made here.
            IsFullPathLink = true
        };

        // Save the workbook as HTML using the configured options.
        workbook.Save(htmlPath, htmlOptions);
    }
}
