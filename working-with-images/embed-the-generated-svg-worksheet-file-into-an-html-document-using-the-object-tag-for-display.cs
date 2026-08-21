// Title: Render Excel worksheet to SVG with Aspose.Cells and embed it using the <object> tag (C#)
// Description: Creates a workbook, fills sample data, converts the first worksheet to an SVG file with FitToViewPort enabled, then generates an HTML page that displays the SVG via an <object> element. Both files are saved to disk.
// Keywords: Aspose.Cells SVG rendering C# | Excel to SVG conversion | embed SVG in HTML object tag | FitToViewPort SvgImageOptions | C# generate worksheet SVG | display Excel data as SVG | Aspose.Cells SheetRender example | HTML page with embedded SVG
// Common Searches: Aspose.Cells render worksheet to SVG C# | how to embed generated SVG in HTML using object tag | C# convert Excel sheet to scalable SVG | display Excel worksheet as SVG in web page | FitToViewPort option Aspose.Cells SVG
// Developer Intent: Produce an SVG representation of an Excel worksheet and show it in a web page via the <object> element.
// Use Cases: Integrate high‑resolution worksheet graphics into dashboards without raster artifacts. | Create printable SVG reports that can be viewed directly in browsers. | Build responsive web pages where Excel data scales smoothly on any device.
// AI Prompts: Generate C# code that uses Aspose.Cells to export a worksheet to SVG with FitToViewPort set to true. | Write an HTML template that embeds a given SVG file using the <object> tag and makes it fill the viewport. | Explain how to modify the example to embed multiple worksheet SVGs on one HTML page, each inside its own <object> element.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgToHtml
{
    // Creates a workbook, fills sample data, converts the first worksheet to an SVG file with FitToViewPort enabled, then generates an HTML page that displays the SVG via an <object> element. Both files are saved to disk.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");

                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(210);
                sheet.Cells["B4"].PutValue(150);

                // Configure SVG rendering options (no ImageFormat property needed)
                SvgImageOptions svgOptions = new SvgImageOptions
                {
                    FitToViewPort = true // Make SVG fit the viewport
                };

                // Render the worksheet to an SVG file
                string svgFileName = "worksheet.svg";
                SheetRender renderer = new SheetRender(sheet, svgOptions);
                renderer.ToImage(0, svgFileName);

                // Build an HTML document that embeds the generated SVG using the <object> tag
                string htmlContent = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset=""UTF-8"">
    <title>Worksheet SVG in HTML</title>
    <style>
        body, html {{ margin:0; padding:0; height:100%; }}
        object {{ width:100%; height:100%; border:none; }}
    </style>
</head>
<body>
    <object data=""{svgFileName}"" type=""image/svg+xml""></object>
</body>
</html>";

                // Save the HTML file
                string htmlFileName = "worksheet.html";
                File.WriteAllText(htmlFileName, htmlContent);

                Console.WriteLine($"SVG file generated: {Path.GetFullPath(svgFileName)}");
                Console.WriteLine($"HTML file generated: {Path.GetFullPath(htmlFileName)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
