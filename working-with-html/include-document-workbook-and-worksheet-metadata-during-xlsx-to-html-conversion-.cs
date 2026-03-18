using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        string sourcePath = "Sample.xlsx";
        Workbook workbook = new Workbook(sourcePath);

        // ----- Document (built‑in) properties -----
        workbook.BuiltInDocumentProperties.Author = "John Doe";
        workbook.BuiltInDocumentProperties.Title = "Sample Workbook";
        workbook.BuiltInDocumentProperties.Company = "Acme Corp";

        // ----- Custom document properties -----
        workbook.CustomDocumentProperties.Add("Project", "AsposeDemo");
        workbook.CustomDocumentProperties.Add("Version", 1.0);

        // ----- Worksheet metadata (example) -----
        Worksheet sheet = workbook.Worksheets[0];
        sheet.TabColor = Color.LightBlue;   // set tab color
        sheet.IsVisible = true;             // ensure worksheet is visible
        sheet.Unprotect();                  // remove protection if any

        // Create HTML save options and enable exporting of all metadata
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            ExportDocumentProperties = true,   // export document properties
            ExportWorkbookProperties = true,   // export workbook properties
            ExportWorksheetProperties = true   // export worksheet properties
        };

        // Save the workbook as HTML, preserving metadata
        string htmlPath = "Sample.html";
        workbook.Save(htmlPath, options);
    }
}