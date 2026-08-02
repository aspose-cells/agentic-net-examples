// Title: Render Excel Worksheet to SVG with Aspose.Cells and Embed in HTML via <object> Tag (C#)
// Description: Creates a workbook, fills cells with sample data, uses Aspose.Cells SheetRender to export the first worksheet as an SVG file, then generates an HTML page that displays the SVG through the <object> element. Includes error handling and file‑system checks.
// Keywords: Aspose.Cells | C# SVG rendering | Excel to SVG conversion | SheetRender | embed SVG in HTML | object tag | web preview of worksheet | display SVG in browser
// Common Searches: Aspose.Cells convert worksheet to SVG C# | How to embed SVG file in HTML using object tag | C# render Excel sheet as SVG | Display Excel worksheet as SVG on a web page | Save worksheet as SVG and generate HTML page
// Developer Intent: Generate an SVG image of an Excel worksheet and embed it in an HTML document using the <object> tag for browser display.
// Use Cases: Provide a lightweight, printable snapshot of a report sheet for web preview without requiring Excel. | Integrate Excel‑derived visualizations into dashboards where the SVG updates automatically when the workbook changes. | Create static HTML reports that include worksheet graphics, enabling cross‑platform viewing in any modern browser.
// AI Prompts: Write C# code that uses Aspose.Cells to render a specific worksheet to an SVG file and embed the result in an HTML page with an <object> element. | Add comprehensive error handling for SVG rendering and HTML file creation when using Aspose.Cells in .NET. | Show how to customize the width, height, and fallback content of the <object> tag for optimal SVG display.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Creates a workbook, fills cells with sample data, uses Aspose.Cells SheetRender to export the first worksheet as an SVG file, then generates an HTML page that displays the SVG through the <object> element. Includes error handling and file‑system checks.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(120);
            worksheet.Cells["A3"].PutValue("Orange");
            worksheet.Cells["B3"].PutValue(85);

            // Set SVG rendering options (default options are sufficient)
            SvgImageOptions svgOptions = new SvgImageOptions();

            string svgFile = "worksheet.svg";

            // Ensure the output directory for the SVG exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(svgFile));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Render the worksheet to SVG
            try
            {
                SheetRender renderer = new SheetRender(worksheet, svgOptions);
                renderer.ToImage(0, svgFile);
            }
            catch (Exception renderEx)
            {
                Console.Error.WriteLine($"SVG rendering error: {renderEx.Message}");
                throw;
            }

            // Build an HTML document that embeds the generated SVG using the <object> tag
            string html = $@"
<html>
<head>
    <title>Worksheet SVG</title>
    <meta charset=""UTF-8"">
</head>
<body>
    <object data=""{svgFile}"" type=""image/svg+xml"" width=""100%"" height=""100%"">
        Your browser does not support SVG.
    </object>
</body>
</html>";

            // Save the HTML file
            string htmlFile = "worksheet.html";
            try
            {
                File.WriteAllText(htmlFile, html);
            }
            catch (Exception ioEx)
            {
                Console.Error.WriteLine($"Failed to write HTML file: {ioEx.Message}");
                throw;
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
