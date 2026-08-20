// Title: C# – Export Excel Worksheet to UTF‑8 SVG using Aspose.Cells
// Description: Demonstrates how to create a workbook, insert multilingual text, configure SvgImageOptions, and render the first worksheet to a UTF‑8 encoded SVG file with SheetRender, then save the workbook as XLSX.
// Keywords: Aspose.Cells C# SVG export | UTF-8 SVG Aspose | SvgImageOptions | SheetRender Unicode | Excel to SVG conversion | multilingual SVG output
// Common Searches: export excel to svg asp.net | aspocells svg utf-8 encoding | c# render worksheet as svg | preserve unicode characters in svg export | sheetrender svg options
// Developer Intent: Create an SVG image of a worksheet that retains all Unicode characters by using UTF‑8 encoding.
// Use Cases: Generate scalable SVG charts for web pages that include Chinese, Russian, or Japanese labels. | Produce high‑resolution documentation screenshots of Excel sheets without losing non‑Latin text. | Automate batch conversion of multiple worksheets to SVG files for archival while ensuring correct character encoding.
// AI Prompts: Write C# code with Aspose.Cells to render a worksheet to a UTF‑8 SVG, including FitToViewPort and a custom CSS prefix. | Explain the role of SvgImageOptions in preserving Unicode text when exporting Excel to SVG with Aspose.Cells. | Provide a step‑by‑step guide to batch‑process several worksheets into SVG files, guaranteeing UTF‑8 encoding for all outputs.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to create a workbook, insert multilingual text, configure SvgImageOptions, and render the first worksheet to a UTF‑8 encoded SVG file with SheetRender, then save the workbook as XLSX.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add Unicode text to demonstrate UTF‑8 preservation in the SVG
            sheet.Cells["A1"].PutValue("中文字符");      // Chinese
            sheet.Cells["A2"].PutValue("Привет");       // Russian
            sheet.Cells["A3"].PutValue("こんにちは");   // Japanese

            // Configure SVG rendering options
            SvgImageOptions svgOptions = new SvgImageOptions
            {
                // Ensure SVG output (default for SvgImageOptions)
                FitToViewPort = true,        // Optional: fit content to viewport
                CssPrefix = "sheet-"          // Optional CSS prefix
            };

            // Render the worksheet to an SVG file (uses SheetRender with SvgImageOptions)
            SheetRender renderer = new SheetRender(sheet, svgOptions);
            renderer.ToImage(0, "worksheet.svg");

            // Save the workbook itself (lifecycle save rule)
            workbook.Save("worksheet.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
