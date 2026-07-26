// Title: Export Aspose.Cells Workbook to HTML with Per‑Worksheet CSS Using a Custom FilePathProvider (C#)
// Description: Creates a workbook with two sheets, enables HtmlSaveOptions.ExportWorksheetCSSSeparately, and implements a custom IFilePathProvider that writes each sheet's CSS file into a "css" subfolder. The resulting HTML references the correct external stylesheet for every worksheet.
// Keywords: Aspose.Cells | C# | HtmlSaveOptions | ExportWorksheetCSSSeparately | IFilePathProvider | per worksheet CSS | HTML export | custom CSS folder | external stylesheet | Aspose.Cells HTML export
// Common Searches: Aspose.Cells export HTML separate CSS per worksheet | How to use ExportWorksheetCSSSeparately in Aspose.Cells | Custom IFilePathProvider example for CSS files | Save workbook as HTML with external CSS files | Aspose.Cells C# generate CSS files in subfolder
// Developer Intent: Generate HTML from a workbook where each worksheet links to its own external CSS file placed in a designated folder.
// Use Cases: Create web‑ready reports with isolated styling for each sheet | Maintain individual CSS files for version control and theming | Automate HTML documentation generation in CI/CD pipelines | Allow front‑end developers to edit sheet‑specific styles without rebuilding the workbook
// AI Prompts: Provide C# code that saves an Aspose.Cells workbook to HTML with separate CSS files for each worksheet and stores them in a 'css' directory. | Explain the role of IFilePathProvider in customizing CSS output paths and show a sample implementation. | Step‑by‑step instructions to configure HtmlSaveOptions for per‑worksheet CSS and verify the HTML links.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook with two sheets, enables HtmlSaveOptions.ExportWorksheetCSSSeparately, and implements a custom IFilePathProvider that writes each sheet's CSS file into a "css" subfolder. The resulting HTML references the correct external stylesheet for every worksheet.
class ExportWorksheetCssSeparatelyDemo
{
    static void Main()
    {
        // Create a new workbook and add sample data to multiple worksheets
        Workbook workbook = new Workbook();

        // First worksheet (default)
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";
        sheet1.Cells["A1"].PutValue("Data in Sheet1");

        // Second worksheet
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
        sheet2.Cells["A1"].PutValue("Data in Sheet2");

        // Configure HTML save options to export CSS for each worksheet separately
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportWorksheetCSSSeparately = true;

        // Use a custom file path provider to place CSS files in a subfolder and name them per worksheet
        saveOptions.FilePathProvider = new CssFilePathProvider("css");

        // Ensure the output directory exists
        string outputDir = Path.Combine(Environment.CurrentDirectory, "HtmlOutput");
        Directory.CreateDirectory(outputDir);

        // Save the workbook as HTML; each worksheet will have its own CSS file referenced correctly
        workbook.Save(Path.Combine(outputDir, "Workbook.html"), saveOptions);
    }
}

// Custom IFilePathProvider implementation to generate CSS file paths per worksheet
class CssFilePathProvider : IFilePathProvider
{
    private readonly string _cssFolder;

    public CssFilePathProvider(string cssFolder)
    {
        _cssFolder = cssFolder;
        // Create the CSS folder if it does not exist
        string fullPath = Path.Combine(Environment.CurrentDirectory, _cssFolder);
        Directory.CreateDirectory(fullPath);
    }

    // Returns the relative path for the CSS file of a given worksheet
    public string GetFullName(string sheetName)
    {
        // Example output: "css/Sheet1.css"
        return Path.Combine(_cssFolder, $"{sheetName}.css");
    }
}
