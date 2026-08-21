// Title: Convert Excel to HTML with PresentationPreference = BestFit using Aspose.Cells for .NET (C#)
// Description: Shows how to export a workbook as HTML with PresentationPreference enabled (BestFit) and without full‑path links, using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | Excel to HTML | PresentationPreference | BestFit | HtmlSaveOptions | IsFullPathLink | workbook export | HTML conversion
// Common Searches: Aspose.Cells enable PresentationPreference BestFit | save Excel as HTML without full path links .NET | HTML conversion of workbook with adjusted layout | C# Aspose.Cells HtmlSaveOptions examples | export Excel to web‑friendly HTML
// Developer Intent: Generate an HTML file from an Excel workbook that automatically fits column widths and row heights, using Aspose.Cells with PresentationPreference set to BestFit.
// Use Cases: Quick preview of spreadsheets on a website with optimal column sizing. | Exporting Excel reports to HTML for embedding in intranet portals or emails. | Creating web‑ready versions of dashboards without exposing local file paths.
// AI Prompts: Write C# code that loads an existing .xlsx file and saves it as HTML with PresentationPreference = BestFit and IsFullPathLink disabled using Aspose.Cells. | Explain the effect of PresentationPreference on column width and row height when converting Excel to HTML with Aspose.Cells. | Provide a step‑by‑step guide to customize styles and images in the HTML output after enabling PresentationPreference in Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to export a workbook as HTML with PresentationPreference enabled (BestFit) and without full‑path links, using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (empty Excel file)
        Workbook workbook = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Presentation Preference Test");
        sheet.Cells["B1"].PutValue(123.45);
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        // Enable PresentationPreference for a more beautiful layout (BestFit)
        htmlOptions.PresentationPreference = true;
        // Optional: avoid using full path links in generated HTML files
        htmlOptions.IsFullPathLink = false;

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
