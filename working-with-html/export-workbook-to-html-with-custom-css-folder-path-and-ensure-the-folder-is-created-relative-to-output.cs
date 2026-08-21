// Title: Export Aspose.Cells Workbook to HTML with a Custom CSS Subfolder and Auto‑Created Directories (C#)
// Description: Demonstrates how to save a workbook as HTML using Aspose.Cells, generate a separate CSS file for each worksheet, place those files in a user‑defined subfolder, and automatically create the output and CSS directories if they do not exist.
// Keywords: Aspose.Cells HTML export C# | custom CSS folder Aspose.Cells | HtmlSaveOptions CreateDirectory | ExportWorksheetCSSSeparately example | IFilePathProvider custom path | save workbook as HTML .NET | auto‑create output folder
// Common Searches: Aspose.Cells export workbook to HTML with custom CSS folder | How to use HtmlSaveOptions CreateDirectory in C# | IFilePathProvider implementation for CSS path Aspose.Cells | Separate CSS files per worksheet Aspose.Cells | C# generate HTML from Excel with custom stylesheet directory
// Developer Intent: Generate an HTML version of an Excel workbook where each worksheet’s stylesheet is saved to a specified subfolder, with all required directories created automatically.
// Use Cases: Create web‑ready reports that keep HTML and CSS files organized in dedicated folders. | Batch‑process multiple workbooks into HTML while maintaining a consistent CSS folder structure. | Integrate HTML export into a web service that serves the generated pages and their styles from a known relative path.
// AI Prompts: Write C# code using Aspose.Cells to export a workbook to HTML, storing each worksheet’s CSS file in a subfolder called "styles" and ensuring the folder is created if missing. | Explain how to implement IFilePathProvider to control CSS file locations when saving a workbook as HTML with Aspose.Cells. | Provide a step‑by‑step guide to configure HtmlSaveOptions with ExportWorksheetCSSSeparately = true and CreateDirectory = true for custom CSS output.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to save a workbook as HTML using Aspose.Cells, generate a separate CSS file for each worksheet, place those files in a user‑defined subfolder, and automatically create the output and CSS directories if they do not exist.
class ExportWorkbookToHtml
{
    static void Main()
    {
        // Create a sample workbook with some data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B2"].PutValue("World");

        // Define the output HTML file path
        string outputDir = Path.Combine(Environment.CurrentDirectory, "HtmlOutput");
        string htmlPath = Path.Combine(outputDir, "sample.html");

        // Define a custom folder (relative to the output directory) for CSS files
        string cssFolder = Path.Combine(outputDir, "custom_css");

        // Ensure the custom CSS folder exists (CreateDirectory will also create the output folder)
        Directory.CreateDirectory(cssFolder);

        // Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // Export CSS for each worksheet as a separate file
            ExportWorksheetCSSSeparately = true,

            // Automatically create missing directories (output folder and CSS folder)
            CreateDirectory = true,

            // Use a custom file path provider to place CSS files into the custom folder
            FilePathProvider = new CustomCssPathProvider("custom_css")
        };

        // Save the workbook as HTML using the configured options
        workbook.Save(htmlPath, saveOptions);

        Console.WriteLine($"HTML file saved to: {htmlPath}");
        Console.WriteLine($"CSS files saved to folder: {cssFolder}");
    }
}

// Custom implementation of IFilePathProvider that directs CSS files to a specific subfolder
class CustomCssPathProvider : IFilePathProvider
{
    private readonly string _relativeCssFolder;

    public CustomCssPathProvider(string relativeCssFolder)
    {
        _relativeCssFolder = relativeCssFolder;
    }

    // Returns a relative path for the CSS file of a given worksheet name
    public string GetFullName(string sheetName)
    {
        // Example: "custom_css/Sheet1.css"
        return Path.Combine(_relativeCssFolder, $"{sheetName}.css");
    }
}
