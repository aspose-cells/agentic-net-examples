// Title: Export RTL Text to HTML with Default HtmlSaveOptions – Aspose.Cells for .NET
// Description: Demonstrates how to enable the worksheet DisplayRightToLeft property, insert Hebrew and Arabic strings, and save the workbook as HTML using a plain HtmlSaveOptions object, preserving right‑to‑left alignment in the generated HTML.
// Keywords: Aspose.Cells | C# | HTML export | right‑to‑left | RTL text | DisplayRightToLeft | Hebrew | Arabic | default HtmlSaveOptions | Excel to HTML conversion
// Common Searches: Aspose.Cells preserve RTL when saving as HTML | default HtmlSaveOptions RTL alignment | export Arabic worksheet to HTML C# | how to keep right‑to‑left direction in HTML output Aspose.Cells | DisplayRightToLeft HTML conversion example
// Developer Intent: Generate an HTML file from a workbook that retains right‑to‑left text direction without custom save settings.
// Use Cases: Create web‑ready reports for Middle‑Eastern users where cells contain Hebrew or Arabic content. | Automate conversion of multilingual Excel files to HTML while maintaining proper text flow. | Produce lightweight HTML previews of RTL spreadsheets for documentation or intranet portals.
// AI Prompts: Provide C# code that saves a worksheet with DisplayRightToLeft enabled to HTML using Aspose.Cells default HtmlSaveOptions. | Show how to validate that the exported HTML contains dir="rtl" attributes for RTL cells. | Explain the impact of the DisplayRightToLeft property on the HTML markup produced by Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to enable the worksheet DisplayRightToLeft property, insert Hebrew and Arabic strings, and save the workbook as HTML using a plain HtmlSaveOptions object, preserving right‑to‑left alignment in the generated HTML.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Enable right‑to‑left display for the worksheet
        worksheet.DisplayRightToLeft = true;

        // Add some right‑to‑left text (Hebrew and Arabic examples)
        worksheet.Cells["A1"].PutValue("שלום עולם");          // Hebrew: "Hello World"
        worksheet.Cells["A2"].PutValue("مرحبا بالعالم");      // Arabic: "Hello World"

        // Use the default HTML save options (no custom settings required)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Save the workbook as HTML; the RTL alignment will be preserved
        workbook.Save("RtlAlignedOutput.html", htmlOptions);
    }
}
