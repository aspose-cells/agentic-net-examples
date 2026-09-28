// Title: Export an Excel worksheet to SVG and embed it in an HTML page using the <object> tag with Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to render a worksheet as a single‑page SVG file and writes an HTML file that displays the SVG via an <object> element. | Modify the example to accept custom width and height values for the <object> tag and make the HTML output responsive. | Add comprehensive error handling that verifies the output directory, catches rendering exceptions, and logs any failures during SVG export or HTML creation.
// Common Searches: how to render an Excel sheet as SVG with Aspose.Cells and show it in a web page | C# embed exported worksheet SVG in HTML using object tag | Aspose.Cells save worksheet to SVG then create HTML viewer | display Excel data as scalable vector graphic in browser with .NET | object tag for embedding SVG generated from Excel workbook
// Tags: Aspose.Cells export worksheet to SVG | C# embed SVG using object tag | HTML viewer for Excel worksheet image | single-page worksheet SVG rendering | responsive SVG embed in web page

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The program creates a workbook, populates cells, saves the first worksheet as a single‑page SVG file using Aspose.Cells, and generates an HTML file that embeds the SVG via an <object> tag for browser display.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Populate the first worksheet with sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);

            // Define output folder and file names
            string outputFolder = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputFolder);
            string svgPath = Path.Combine(outputFolder, "Sheet1.svg");
            string htmlPath = Path.Combine(outputFolder, "index.html");

            // Save the worksheet as an SVG image
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // ImageFormat is inferred from the file extension (SVG)
                OnePagePerSheet = true
            };
            SheetRender sheetRender = new SheetRender(sheet, imgOptions);
            // Render the first (and only) page of the sheet to SVG
            sheetRender.ToImage(0, svgPath);

            // Build HTML content that embeds the SVG using an <object> tag
            string htmlContent = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset=""UTF-8"">
    <title>Worksheet SVG Display</title>
</head>
<body>
    <h1>Worksheet as SVG</h1>
    <object type=""image/svg+xml"" data=""{Path.GetFileName(svgPath)}"" width=""100%"" height=""600px"">
        Your browser does not support SVG.
    </object>
</body>
</html>";

            // Write the HTML file to disk (same folder as the SVG)
            File.WriteAllText(htmlPath, htmlContent);

            Console.WriteLine($"HTML file generated at: {htmlPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
