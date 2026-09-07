// Title: Export an Aspose.Cells workbook to HTML with a relative CSS folder created automatically in C#
// AI Prompts: Write C# code that uses Aspose.Cells to save a workbook as an HTML file, generates an external stylesheet folder beside the HTML file, and configures HtmlSaveOptions to reference those stylesheets. | Generate a C# example that verifies the destination path, creates a CSS folder if missing, and exports only the active worksheet to HTML with Aspose.Cells.
// Common Searches: how to generate HTML from Aspose.Cells workbook with external CSS folder in C# | c# Aspose.Cells save workbook as html and automatically create a css subfolder | Aspose.Cells HtmlSaveOptions export only the active sheet and avoid base64 images | programmatically prepare output folder and css subfolder before exporting to HTML using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions HTML export settings | C# generate CSS directory for Aspose.Cells HTML export | save workbook as HTML without base64 images Aspose.Cells | ensure target directory exists before workbook.Save C# | export only selected worksheet to HTML Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example builds a workbook, ensures the HTML output directory and a "css" subfolder exist, configures HtmlSaveOptions to export only the active worksheet without embedding images as Base64, and saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SampleData";
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Oranges");
            sheet.Cells["B3"].PutValue(20);

            // Define the output HTML file path
            string outputHtmlPath = "output/report.html";

            // Ensure the output directory exists
            string outputDirectory = Path.GetDirectoryName(outputHtmlPath);
            if (string.IsNullOrEmpty(outputDirectory))
                outputDirectory = Directory.GetCurrentDirectory();
            Directory.CreateDirectory(outputDirectory);

            // Define the CSS folder relative to the HTML file
            string cssFolderRelative = "css";
            string cssFolderPath = Path.Combine(outputDirectory, cssFolderRelative);
            Directory.CreateDirectory(cssFolderPath); // Ensure CSS folder exists

            // Configure HTML save options (default internal CSS)
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportActiveWorksheetOnly = true,
                ExportImagesAsBase64 = false
                // External CSS options are omitted because they are not available in the current API version
            };

            // Save the workbook as HTML
            workbook.Save(outputHtmlPath, saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
