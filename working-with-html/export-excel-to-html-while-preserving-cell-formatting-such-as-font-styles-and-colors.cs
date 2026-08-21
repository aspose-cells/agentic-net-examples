// Title: Export Excel to HTML with Full Font Styling using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, apply custom font families, sizes, colors, bold and italic styles, configure HtmlSaveOptions (ExportDataOptions.All, ExcludeUnusedStyles = false) and save the file as HTML while preserving every cell's formatting.
// Keywords: Aspose.Cells HTML export | preserve Excel font styles | C# convert Excel to HTML | HtmlSaveOptions formatting | ExportDataOptions.All example | retain cell colors Aspose .NET | Excel to HTML with CSS | styled HTML report from workbook
// Common Searches: Aspose.Cells keep font color when exporting to HTML | C# export Excel to HTML preserving bold and italic | HtmlSaveOptions retain all styles Aspose | How to save Excel as styled HTML using .NET | Export Excel worksheet to HTML with original formatting
// Developer Intent: Generate an HTML file from an Excel workbook that maintains all font attributes (family, size, color, bold, italic) exactly as they appear in the source sheet.
// Use Cases: Publish a spreadsheet as a web‑ready report that looks identical to the Excel version. | Embed formatted spreadsheet data in emails or documentation without losing visual fidelity. | Provide an on‑the‑fly HTML preview of Excel sheets in a web application, ensuring users see the same styling as in the original file.
// AI Prompts: Show how to embed CSS inline instead of external files when exporting Excel to HTML with Aspose.Cells. | Provide code to export only a selected worksheet to HTML while keeping its cell formatting. | Explain how to disable CSS generation and output pure HTML with inline style attributes using Aspose.Cells.

using System.Drawing;
using Aspose.Cells;

// Demonstrates how to create a workbook, apply custom font families, sizes, colors, bold and italic styles, configure HtmlSaveOptions (ExportDataOptions.All, ExcludeUnusedStyles = false) and save the file as HTML while preserving every cell's formatting.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Cell A1 with custom font style and color
        Cell cellA1 = sheet.Cells["A1"];
        cellA1.PutValue("Hello World");
        Style styleA1 = cellA1.GetStyle();
        styleA1.Font.Name = "Arial";
        styleA1.Font.Size = 12;
        styleA1.Font.Color = Color.Blue;
        styleA1.Font.IsBold = true;
        cellA1.SetStyle(styleA1);

        // Cell B2 with a different font style and color
        Cell cellB2 = sheet.Cells["B2"];
        cellB2.PutValue("Aspose.Cells");
        Style styleB2 = cellB2.GetStyle();
        styleB2.Font.Name = "Times New Roman";
        styleB2.Font.Size = 14;
        styleB2.Font.Color = Color.Green;
        styleB2.Font.IsItalic = true;
        cellB2.SetStyle(styleB2);

        // Configure HTML save options to preserve all formatting
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportDataOptions = HtmlExportDataOptions.All; // export all data
        htmlOptions.ExcludeUnusedStyles = false; // keep all style definitions
        // Additional options can be set as needed, e.g., htmlOptions.DisableCss = false;

        // Save the workbook as an HTML file with the specified options
        workbook.Save("ExportedWorkbook.html", htmlOptions);
    }
}
