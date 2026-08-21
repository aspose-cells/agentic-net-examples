// Title: C# – Convert Excel to Compact HTML with Aspose.Cells PresentationPreference AutoFit
// Description: Loads an Excel workbook, configures HtmlSaveOptions with PresentationPreference set to true for auto‑fit column widths, and saves the file as a space‑efficient HTML document using Aspose.Cells.
// Keywords: Aspose.Cells HTML conversion C# | HtmlSaveOptions PresentationPreference | auto fit columns Aspose.Cells | compact HTML output from Excel | C# save workbook as HTML | PresentationPreference true Aspose | Excel to HTML auto‑fit columns | Aspose.Cells HTML save options
// Common Searches: How to enable PresentationPreference in Aspose.Cells C# | Aspose.Cells convert Excel to HTML with auto‑fit columns | C# HtmlSaveOptions PresentationPreference example | Compact HTML output from Excel using Aspose.Cells
// Developer Intent: Generate HTML from an Excel workbook with auto‑fit column widths for a compact layout using Aspose.Cells in C#.
// Use Cases: Embedding spreadsheet data in web pages with minimal horizontal scrolling. | Creating email‑friendly HTML reports from Excel files. | Producing printable HTML previews that conserve screen space. | Building lightweight HTML dashboards from financial or analytical workbooks.
// AI Prompts: Show how to disable PresentationPreference to keep original column widths when saving to HTML with Aspose.Cells. | Provide C# code that converts an Excel workbook to HTML and applies a custom CSS stylesheet using Aspose.Cells. | Explain the impact of PresentationPreference on column width calculation in Aspose.Cells HTML export. | Give an example of using HtmlSaveOptions to embed a custom JavaScript file in the generated HTML.

using System;
using Aspose.Cells;

// Loads an Excel workbook, configures HtmlSaveOptions with PresentationPreference set to true for auto‑fit column widths, and saves the file as a space‑efficient HTML document using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load the source Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options:
        // PresentationPreference = true enables a more compact, auto‑fit style presentation.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.PresentationPreference = true;

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
