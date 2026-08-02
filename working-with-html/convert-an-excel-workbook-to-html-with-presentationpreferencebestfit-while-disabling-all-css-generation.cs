// Title: C# – Export Excel to HTML with BestFit layout and no CSS using Aspose.Cells
// Description: Learn how to convert an Excel workbook to HTML in C# with Aspose.Cells, enabling the BestFit presentation preference for optimal layout while disabling all CSS generation so only inline styles are emitted.
// Keywords: Aspose.Cells HTML export | C# Excel to HTML conversion | PresentationPreference BestFit | DisableCss Aspose.Cells | inline styles HTML output | .NET Excel to HTML | export Excel without external CSS
// Common Searches: Aspose.Cells export Excel to HTML BestFit | C# disable CSS when saving workbook as HTML | How to get inline‑style HTML from Excel using Aspose | PresentationPreference option in HtmlSaveOptions | Convert Excel file to HTML without stylesheet
// Developer Intent: Generate an HTML version of an Excel workbook in C# with a layout that mimics Excel’s appearance (BestFit) while suppressing external CSS files, relying solely on inline styling.
// Use Cases: Embedding spreadsheet previews in email bodies where external CSS is blocked. | Creating lightweight, self‑contained HTML reports for documentation or archiving. | Building web pages that display Excel data without loading additional stylesheet resources.
// AI Prompts: Write C# code that uses Aspose.Cells to save an Excel workbook as HTML with PresentationPreference set to BestFit and CSS generation disabled. | Explain how PresentationPreference influences the HTML layout and why disabling CSS results in only inline styles. | Provide a step‑by‑step tutorial for configuring HtmlSaveOptions to achieve a best‑fit layout with no external CSS in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Learn how to convert an Excel workbook to HTML in C# with Aspose.Cells, enabling the BestFit presentation preference for optimal layout while disabling all CSS generation so only inline styles are emitted.
class Program
{
    static void Main()
    {
        // Load the source Excel workbook (replace with your actual file path)
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Create HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Enable presentation preference for a more beautiful layout (BestFit)
        htmlOptions.PresentationPreference = true;

        // Disable all CSS generation – only inline styles will be used
        htmlOptions.DisableCss = true;

        // Save the workbook as an HTML file
        string outputFile = "output.html";
        workbook.Save(outputFile, htmlOptions);
    }
}
