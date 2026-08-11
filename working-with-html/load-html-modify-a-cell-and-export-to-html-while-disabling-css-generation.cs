// Title: Load an HTML workbook, change a cell, and export to HTML with only inline styling using Aspose.Cells for .NET
// Description: This C# example shows how to open a workbook from an existing HTML file, modify cell A1 in the first worksheet, set HtmlSaveOptions to suppress external style sheets, and save the workbook back to HTML with all formatting embedded directly in the markup.
// Keywords: Aspose.Cells load HTML C# | modify worksheet cell Aspose.Cells | HtmlSaveOptions DisableCss | export HTML with embedded styles | C# Aspose.Cells inline styling | save workbook as HTML without external CSS | Aspose.Cells HTML to Excel conversion
// Common Searches: Aspose.Cells disable external CSS when saving HTML | How to edit a cell in an HTML workbook using C# | Save Aspose.Cells workbook as HTML with inline styles only | C# code to load HTML file into Aspose.Cells and modify data | Aspose.Cells HtmlSaveOptions examples
// Developer Intent: Open an HTML‑based workbook, update a specific cell value, and re‑export it as HTML while preventing the generation of separate CSS files.
// Use Cases: Refresh data in an HTML spreadsheet without altering its visual layout, keeping all styling inside the file for easy deployment. | Create email‑ready HTML reports where only inline formatting is supported, ensuring consistent rendering across mail clients. | Automate batch processing of HTML workbooks to apply data corrections and produce self‑contained HTML files for web publishing.
// AI Prompts: Generate C# code that loads an HTML file into Aspose.Cells, updates cell B2, and saves the workbook as HTML with all styles inlined. | Explain the effect of HtmlSaveOptions.DisableCss in Aspose.Cells and how it influences the output HTML file. | Provide a script to iterate over a folder of HTML workbooks, change a designated cell in each, and export them with no external CSS using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// This C# example shows how to open a workbook from an existing HTML file, modify cell A1 in the first worksheet, set HtmlSaveOptions to suppress external style sheets, and save the workbook back to HTML with all formatting embedded directly in the markup.
class Program
{
    static void Main()
    {
        // Load the workbook from an existing HTML file
        Workbook workbook = new Workbook("input.html");

        // Modify a specific cell (e.g., A1) in the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Modified Value");

        // Configure HTML save options to use only inline styles (disable external CSS)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.DisableCss = true; // Apply rule: only inline styles

        // Save the modified workbook back to HTML
        workbook.Save("output.html", htmlOptions);
    }
}
