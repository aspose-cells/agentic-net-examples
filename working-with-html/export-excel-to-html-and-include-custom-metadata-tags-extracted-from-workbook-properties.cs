// Title: Export Excel to a Single HTML File with Built‑in and Custom <meta> Tags using Aspose.Cells for .NET (C#)
// Description: Loads an XLSX workbook, sets built‑in properties (Author, Title), adds custom properties (Project, Version, Reviewed) via WorkbookMetadata, saves the metadata, then exports the workbook to one HTML file with all properties embedded as <meta> tags using HtmlSaveOptions (ExportWorkbookProperties & ExportDocumentProperties).
// Keywords: Aspose.Cells export Excel to HTML | C# Excel to HTML with metadata | custom document properties Aspose.Cells | HtmlSaveOptions ExportWorkbookProperties | ExportDocumentProperties HTML | single HTML file Aspose.Cells | SEO meta tags from Excel | GitHub Aspose.Cells example | USA .NET developer guide | AEO Aspose.Cells HTML export
// Common Searches: how to embed workbook properties as meta tags when saving Excel as HTML using Aspose.Cells | Aspose.Cells .NET export Excel to single HTML file with custom metadata | C# example for adding custom document properties and exporting to HTML | Aspose.Cells HtmlSaveOptions ExportWorkbookProperties usage | include SEO meta tags from Excel workbook in HTML output
// Developer Intent: Generate an HTML version of an Excel workbook that retains both built‑in and custom document properties as <meta> tags for downstream processing and SEO.
// Use Cases: Publish a web‑ready report where author, title, project, and version information are searchable by search engines. | Provide compliance‑friendly HTML files that carry version and review status without exposing the original workbook. | Integrate a single‑file HTML export into intranet portals while preserving all workbook metadata for auditing.
// AI Prompts: Show how to export only selected custom properties as <meta> tags in the HTML output. | Give a C# snippet that reads the <meta> tags from the generated HTML file. | Explain how to disable built‑in properties while keeping custom properties in HtmlSaveOptions.

using System;
using Aspose.Cells;
using Aspose.Cells.Metadata;

// Loads an XLSX workbook, sets built‑in properties (Author, Title), adds custom properties (Project, Version, Reviewed) via WorkbookMetadata, saves the metadata, then exports the workbook to one HTML file with all properties embedded as <meta> tags using HtmlSaveOptions (ExportWorkbookProperties & ExportDocumentProperties).
class ExportExcelToHtmlWithMetadata
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(sourcePath);

        // Set built‑in document properties
        workbook.BuiltInDocumentProperties.Author = "John Doe";
        workbook.BuiltInDocumentProperties.Title = "Sample Report";

        // Add custom document properties via WorkbookMetadata
        WorkbookMetadata metadata = new WorkbookMetadata(
            sourcePath,
            new MetadataOptions(MetadataType.DocumentProperties));

        metadata.CustomDocumentProperties.Add("Project", "Alpha");
        metadata.CustomDocumentProperties.Add("Version", 2);
        metadata.CustomDocumentProperties.Add("Reviewed", true);

        // Save the metadata back to the workbook file
        metadata.Save(sourcePath);

        // Reload the workbook to ensure custom properties are attached
        workbook = new Workbook(sourcePath);

        // Configure HTML save options to export workbook and document properties as <meta> tags
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportWorkbookProperties = true;   // embed built‑in and custom properties
        htmlOptions.ExportDocumentProperties = true; // also export document properties
        htmlOptions.SaveAsSingleFile = true;          // generate a single HTML file

        // Save the workbook as HTML
        string htmlPath = "output.html";
        workbook.Save(htmlPath, htmlOptions);
    }
}
