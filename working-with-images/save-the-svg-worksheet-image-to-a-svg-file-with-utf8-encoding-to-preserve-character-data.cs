// Title: Save Aspose.Cells Worksheet as UTF‑8 SVG in C# – Preserve Unicode Characters
// Description: Demonstrates how to render an Aspose.Cells worksheet containing Chinese, Russian, and Japanese text to SVG, convert the output to a UTF‑8 string, and write it to a file without a BOM, ensuring all Unicode characters are retained.
// Keywords: Aspose.Cells | C# SVG export | UTF-8 SVG | Unicode Excel to SVG | SvgImageOptions | SheetRender | WriteAllText UTF8Encoding | multilingual Excel SVG | no BOM
// Common Searches: C# save worksheet as SVG UTF-8 | Aspose.Cells export SVG Unicode | write SVG file without BOM .NET | render Excel sheet to SVG preserving characters | SvgImageOptions ImageType.Svg example
// Developer Intent: Export a worksheet to an SVG file encoded in UTF‑8 (without BOM) so that all Unicode text displays correctly.
// Use Cases: Generate web‑ready SVG images of multilingual Excel sheets while keeping Chinese, Russian, and Japanese characters readable. | Create SVG reports from Excel data that must be UTF‑8 encoded for downstream processing pipelines. | Capture SVG output in a memory stream, convert it to a UTF‑8 string, and save it with custom encoding settings.
// AI Prompts: Write C# code using Aspose.Cells to render a worksheet with Unicode text to an SVG file encoded in UTF‑8 without a BOM. | Show how to configure SvgImageOptions and SheetRender to export an Excel sheet to SVG while preserving multilingual characters. | Provide a snippet that reads SVG bytes from a MemoryStream, converts them to a UTF‑8 string, and saves the result using File.WriteAllText with UTF8Encoding(false).

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Demonstrates how to render an Aspose.Cells worksheet containing Chinese, Russian, and Japanese text to SVG, convert the output to a UTF‑8 string, and write it to a file without a BOM, ensuring all Unicode characters are retained.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add Unicode text to demonstrate UTF‑8 preservation
        worksheet.Cells["A1"].PutValue("中文字符");   // Chinese
        worksheet.Cells["A2"].PutValue("Привет");    // Russian
        worksheet.Cells["A3"].PutValue("こんにちは"); // Japanese

        // Configure SVG rendering options
        SvgImageOptions svgOptions = new SvgImageOptions
        {
            ImageType = ImageType.Svg,   // Ensure SVG output
            FitToViewPort = true        // Optional: fit content to viewport
        };

        // Render the worksheet to SVG in a memory stream
        SheetRender renderer = new SheetRender(worksheet, svgOptions);
        using (MemoryStream svgStream = new MemoryStream())
        {
            renderer.ToImage(0, svgStream);   // Render first (and only) sheet
            svgStream.Position = 0;           // Reset for reading

            // Convert the stream bytes to a UTF‑8 string
            string svgContent = Encoding.UTF8.GetString(svgStream.ToArray());

            // Write the SVG content to a file using UTF‑8 encoding (no BOM)
            File.WriteAllText("Worksheet.svg", svgContent, new UTF8Encoding(false));
        }

        Console.WriteLine("Worksheet saved as SVG with UTF‑8 encoding.");
    }
}
