using System;
using Aspose.Cells;

class IncludeMetadataInHtml
{
    static void Main()
    {
        // Load the source XLSX file
        Workbook workbook = new Workbook("input.xlsx");

        // Set built‑in document properties
        workbook.BuiltInDocumentProperties.Author = "John Doe";
        workbook.BuiltInDocumentProperties.Title = "Quarterly Sales Report";
        workbook.BuiltInDocumentProperties.Subject = "Sales Data";

        // Add a custom document property
        workbook.CustomDocumentProperties.Add("ProjectId", 12345);

        // Configure HTML save options to include metadata in the output
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        // Export built‑in document properties (default is true, set explicitly for clarity)
        htmlOptions.ExportDocumentProperties = true;
        // Export workbook properties such as author, title, etc.
        htmlOptions.ExportWorkbookProperties = true;
        // Export worksheet properties if needed
        htmlOptions.ExportWorksheetProperties = true;

        // Save the workbook as HTML with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}