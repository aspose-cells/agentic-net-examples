using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and populate it with sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Price");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(1.2);
        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["B3"].PutValue(0.8);

        // Configure SVG rendering options
        SvgImageOptions svgOptions = new SvgImageOptions();
        svgOptions.FitToViewPort = true;          // Fit SVG to viewport
        svgOptions.IsOptimized = true;            // Enable built‑in optimizations
        svgOptions.ImageType = Aspose.Cells.Drawing.ImageType.Svg;

        // Render the worksheet to SVG and capture the output in memory
        SheetRender renderer = new SheetRender(worksheet, svgOptions);
        using (MemoryStream svgStream = new MemoryStream())
        {
            renderer.ToImage(0, svgStream);
            svgStream.Position = 0;
            string svgContent = new StreamReader(svgStream).ReadToEnd();

            // Minify the SVG by removing comments and unnecessary whitespace
            string minifiedSvg = MinifySvg(svgContent);

            // Save the minified SVG to a file
            File.WriteAllText("MinifiedWorksheet.svg", minifiedSvg);
        }

        Console.WriteLine("Minified SVG file has been saved.");
    }

    // Helper method that removes XML comments and collapses whitespace between tags
    static string MinifySvg(string svg)
    {
        // Remove XML comments (<!-- ... -->)
        string withoutComments = Regex.Replace(svg, @"<!--(.*?)-->", string.Empty, RegexOptions.Singleline);

        // Collapse whitespace between tags (e.g., >   < becomes ><)
        string collapsed = Regex.Replace(withoutComments, @">\s+<", "><");

        // Trim leading and trailing whitespace
        return collapsed.Trim();
    }
}