// Title: Create an SVG file from an Aspose.Cells worksheet and embed it in an HTML page using the <object> tag (C#)
// AI Prompts: Write C# code that converts the first worksheet of a new Aspose.Cells workbook to an SVG file and generates an HTML file that embeds the SVG with an <object> element. | Adjust the generated HTML so the <object> tag includes width and height attributes that scale responsively with the browser window. | Add logic to verify that the SVG file was successfully created before writing the HTML file, and log an error if the file is missing.
// Common Searches: C# Aspose.Cells export worksheet to SVG and embed in HTML page | How to display an SVG generated from Excel using the <object> tag in ASP.NET | Responsive HTML layout for SVG output from Aspose.Cells workbook | Validate SVG file existence before creating HTML with Aspose.Cells C# | One-page-per-sheet SVG rendering with Aspose.Cells and HTML embedding
// Tags: Aspose.Cells worksheet to SVG conversion | C# embed SVG using object tag | generate HTML with embedded SVG from Excel | responsive SVG display in HTML | check SVG file existence Aspose.Cells | one-page-per-sheet SVG rendering

using Aspose.Cells;
using Aspose.Cells.Rendering;
using System;
using System.IO;

// The program creates a new Aspose.Cells workbook, fills the first worksheet with sample data, renders that sheet to a single-page SVG file, and writes an HTML document that embeds the SVG via an <object> element for browser display.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Populate the first worksheet with sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["B1"].PutValue("Data");
            sheet.Cells["A2"].PutValue(100);
            sheet.Cells["B2"].PutValue(200);

            // Define file paths for the SVG and the resulting HTML
            string svgFile = "worksheet.svg";
            string htmlFile = "worksheet.html";

            // Save the worksheet as an SVG image (lifecycle rule: save)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                SaveFormat = SaveFormat.Svg,
                OnePagePerSheet = true
            };
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            renderer.ToImage(0, svgFile); // Generates the SVG file

            // Build an HTML document that embeds the SVG using the <object> tag
            string htmlContent = $"<!DOCTYPE html>{Environment.NewLine}" +
                                 $"<html>{Environment.NewLine}" +
                                 $"<head>{Environment.NewLine}" +
                                 $"    <meta charset=\"UTF-8\">{Environment.NewLine}" +
                                 $"    <title>Worksheet SVG</title>{Environment.NewLine}" +
                                 $"</head>{Environment.NewLine}" +
                                 $"<body>{Environment.NewLine}" +
                                 $"    <object data=\"{svgFile}\" type=\"image/svg+xml\" width=\"100%\" height=\"100%\">{Environment.NewLine}" +
                                 $"        Your browser does not support SVG.{Environment.NewLine}" +
                                 $"    </object>{Environment.NewLine}" +
                                 $"</body>{Environment.NewLine}" +
                                 $"</html>";

            // Write the HTML content to a file (lifecycle rule: save)
            File.WriteAllText(htmlFile, htmlContent);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
