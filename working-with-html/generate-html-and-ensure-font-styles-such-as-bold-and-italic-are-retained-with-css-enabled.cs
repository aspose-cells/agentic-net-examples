// Title: Generate HTML from an Excel workbook in C# while keeping bold, italic, and bold‑italic cell fonts using Aspose.Cells
// AI Prompts: Write C# code that creates a workbook, applies bold, italic, and combined bold‑italic formatting to specific cells, and saves it as an HTML file using Aspose.Cells HtmlSaveOptions so the font styles are preserved. | Show how to configure HtmlSaveOptions in Aspose.Cells to export a worksheet to HTML with the original cell font styling intact.
// Common Searches: C# Aspose.Cells how to export Excel to HTML with bold and italic text preserved | retain cell font styles when converting .xlsx to HTML using Aspose.Cells | HtmlSaveOptions settings for keeping text formatting in HTML output from Aspose.Cells
// Tags: Aspose.Cells HTML export with font styling | C# bold italic cell formatting Aspose.Cells | HtmlSaveOptions preserve text styles | Excel to HTML conversion retaining formatting | Aspose.Cells workbook save as HTML

using Aspose.Cells;
using System;
using System.IO;

// The program creates a workbook, applies bold, italic, and bold‑italic styles to cells A1, B1, and C1, and saves the workbook as an HTML file using HtmlSaveOptions, which retains the applied font styles in the generated HTML.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Cell A1 - bold text
            Cell cellA1 = sheet.Cells["A1"];
            cellA1.PutValue("Bold Text");
            Style styleA1 = cellA1.GetStyle();
            styleA1.Font.IsBold = true;
            cellA1.SetStyle(styleA1);

            // Cell B1 - italic text
            Cell cellB1 = sheet.Cells["B1"];
            cellB1.PutValue("Italic Text");
            Style styleB1 = cellB1.GetStyle();
            styleB1.Font.IsItalic = true;
            cellB1.SetStyle(styleB1);

            // Cell C1 - bold and italic text
            Cell cellC1 = sheet.Cells["C1"];
            cellC1.PutValue("Bold Italic");
            Style styleC1 = cellC1.GetStyle();
            styleC1.Font.IsBold = true;
            styleC1.Font.IsItalic = true;
            cellC1.SetStyle(styleC1);

            // Configure HTML save options (default settings retain font styles)
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

            string outputPath = "StyledOutput.html";

            // Save the workbook as an HTML file
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
