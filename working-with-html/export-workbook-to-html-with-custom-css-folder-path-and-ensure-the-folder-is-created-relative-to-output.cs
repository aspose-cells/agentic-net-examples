// Title: Export Aspose.Cells Workbook to HTML with Custom CSS Folder (C#)
// Description: Creates a workbook, defines an output directory and a subfolder for CSS, ensures the folder exists, configures HtmlSaveOptions to export worksheet CSS separately, sets AttachedFilesDirectory to the custom path, enables automatic directory creation, and saves the workbook as HTML.
// Keywords: Aspose.Cells | C# | HTML export | HtmlSaveOptions | ExportWorksheetCSSSeparately | AttachedFilesDirectory | custom CSS folder | CreateDirectory | separate CSS files | relative path
// Common Searches: Aspose.Cells export workbook to html custom css folder | HtmlSaveOptions set attached files directory C# | Create directory for html export Aspose.Cells | Export worksheet CSS separately .NET | Save workbook as html with separate css files
// Developer Intent: Generate an HTML file from a workbook and automatically place the generated CSS files into a user‑specified subfolder relative to the output location.
// Use Cases: Publish web‑ready reports where HTML and style sheets are organized in distinct folders for straightforward deployment. | Batch‑convert many workbooks to HTML while maintaining a consistent CSS directory structure across all outputs. | Integrate HTML conversion into CI/CD pipelines that require on‑the‑fly creation of the output folder hierarchy.
// AI Prompts: Show how to accept a user‑provided CSS folder path, validate it, and ensure the directory exists before saving the HTML. | Provide a C# example that saves the workbook as HTML with images in an "images" subfolder and CSS in a "styles" subfolder using Aspose.Cells. | Explain how to configure HtmlSaveOptions to compress generated CSS files into a zip archive during HTML export. | Demonstrate how to export multiple worksheets to separate HTML files while sharing a common custom CSS folder.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, defines an output directory and a subfolder for CSS, ensures the folder exists, configures HtmlSaveOptions to export worksheet CSS separately, sets AttachedFilesDirectory to the custom path, enables automatic directory creation, and saves the workbook as HTML.
class ExportWorkbookToHtml
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello World");
        sheet.Cells["B2"].PutValue(12345);

        // Define the output directory for the HTML file
        string outputDir = Path.Combine(Environment.CurrentDirectory, "HtmlOutput");
        // Define a relative folder name where the CSS files will be placed
        string cssFolderName = "custom_css";
        // Combine to get the full path of the CSS folder (relative to the output directory)
        string cssFolderPath = Path.Combine(outputDir, cssFolderName);

        // Ensure the CSS folder exists; CreateDirectory will also create any missing parent folders
        Directory.CreateDirectory(cssFolderPath);

        // Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        // Export the worksheet CSS as separate files
        saveOptions.ExportWorksheetCSSSeparately = true;
        // Set the directory where attached files (including the separate CSS) will be saved
        saveOptions.AttachedFilesDirectory = cssFolderPath;
        // Let Aspose.Cells automatically create any missing directories during save
        saveOptions.CreateDirectory = true;

        // Define the full path for the resulting HTML file
        string htmlFilePath = Path.Combine(outputDir, "workbook.html");

        // Save the workbook as HTML using the configured options
        workbook.Save(htmlFilePath, saveOptions);

        Console.WriteLine($"HTML file saved to: {htmlFilePath}");
        Console.WriteLine($"CSS files saved to: {cssFolderPath}");
    }
}
