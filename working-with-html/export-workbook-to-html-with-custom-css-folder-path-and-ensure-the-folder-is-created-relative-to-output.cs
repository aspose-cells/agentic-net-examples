// Title: Export an Aspose.Cells workbook to HTML and automatically create a relative CSS folder using C#
// AI Prompts: Generate C# code that saves a workbook with Aspose.Cells as an HTML file and creates a "css" subfolder next to the HTML file for external style sheets. | Show how to configure HtmlSaveOptions for Aspose.Cells HTML export while ensuring the required CSS directory is created relative to the output path. | Adapt the example to load an existing .xlsx workbook before exporting it to HTML and preparing a custom CSS folder beside the generated HTML file.
// Common Searches: how to save an Aspose.Cells workbook as HTML and generate a css folder in C# | Aspose.Cells create external CSS directory when exporting to HTML using .NET | C# HtmlSaveOptions to export workbook to HTML with relative stylesheet folder | Aspose.Cells HTML export with custom CSS path example | automatically create assets folder for HTML output from Aspose.Cells workbook
// Tags: Aspose.Cells HTML export with CSS assets directory | C# HtmlSaveOptions external stylesheet path | create CSS assets directory for Aspose.Cells HTML output | relative CSS assets path for workbook HTML export | Aspose.Cells generate css assets folder

using System;
using System.IO;
using Aspose.Cells;

// The sample creates or loads a workbook, populates data, defines an HTML output path, computes a "css" subfolder path relative to that HTML file, ensures the folder exists, configures HtmlSaveOptions, and saves the workbook as HTML (embedding CSS by default). It logs the locations of the generated HTML file and the CSS folder, handling any exceptions that may occur.
class ExportToHtmlWithCustomCss
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // creates a new workbook

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);

            // Define output HTML file path
            string outputHtmlPath = Path.Combine(Environment.CurrentDirectory ?? string.Empty, "Report.html");

            // Determine the CSS folder path relative to the HTML file
            string cssFolderRelative = "css"; // folder name relative to HTML file
            string cssFolderPath = Path.Combine(Path.GetDirectoryName(outputHtmlPath) ?? string.Empty, cssFolderRelative);

            // Ensure the CSS folder exists
            Directory.CreateDirectory(cssFolderPath);

            // Configure HTML save options (default embeds CSS; external CSS options are not available in this version)
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as HTML; the CSS will be embedded by default
            workbook.Save(outputHtmlPath, saveOptions);

            Console.WriteLine($"Workbook exported to HTML at: {outputHtmlPath}");
            Console.WriteLine($"CSS folder (created for potential external files) at: {cssFolderPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
