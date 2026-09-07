// Title: Export an Aspose.Cells workbook to HTML with timestamped filenames using a custom IFilePathProvider in C#
// AI Prompts: Write C# code that saves a Workbook as HTML with Aspose.Cells, configuring HtmlSaveOptions to use a custom IFilePathProvider that appends a yyyyMMdd_HHmmss timestamp to each generated file. | Demonstrate how to implement IFilePathProvider.GetFilePath to include an optional index suffix and a formatted timestamp for HTML export resource files. | Modify the example to prepend a user-defined prefix to the HTML file name while still applying the timestamp logic in the custom file path provider.
// Common Searches: Aspose.Cells C# how to add timestamps to HTML export filenames | custom IFilePathProvider example for timestamped HTML files in Aspose.Cells | save Excel workbook as HTML with indexed image files using Aspose.Cells | C# HtmlSaveOptions file naming pattern with date and time stamp | Aspose.Cells export to HTML with custom file path provider and timestamp
// Tags: Aspose.Cells HtmlSaveOptions custom file naming | timestamped filenames IFilePathProvider C# | HTML export indexed resource files Aspose.Cells | C# workbook to HTML with date-time suffix | custom file path provider for Aspose.Cells HTML export

using System;
using System.IO;
using Aspose.Cells;

// The example implements a CustomFilePathProvider that adds a yyyyMMdd_HHmmss timestamp (and optional index) to filenames generated during HTML export. The provider is assigned to HtmlSaveOptions, a sample workbook is created, and workbook.Save writes the HTML and its resources using the timestamped naming pattern.
class CustomFilePathProvider : IFilePathProvider
{
    // Generates a file path that includes a timestamp and optional index.
    public string GetFilePath(string originalFilePath, string fileExtension, int index)
    {
        // Timestamp format: yyyyMMdd_HHmmss
        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

        // Extract directory and base file name without extension.
        string directory = Path.GetDirectoryName(originalFilePath);
        string baseName = Path.GetFileNameWithoutExtension(originalFilePath);

        // Include index for multiple files (e.g., images) if needed.
        string indexedPart = index > 0 ? $"_{index}" : string.Empty;

        // Build new file name with timestamp.
        string newFileName = $"{baseName}_{timestamp}{indexedPart}{fileExtension}";

        // Combine directory and new file name.
        return Path.Combine(directory, newFileName);
    }

    // Required by IFilePathProvider – returns the full name for the main HTML file.
    public string GetFullName(string originalFilePath)
    {
        // No custom modification needed for the main file; return as‑is.
        return originalFilePath;
    }
}

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one).
            Workbook workbook = new Workbook();

            // Populate the workbook with sample data.
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["B1"].PutValue("Data");
            sheet.Cells["A2"].PutValue(100);
            sheet.Cells["B2"].PutValue(200);

            // Configure HTML save options with the custom file path provider.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            htmlOptions.FilePathProvider = new CustomFilePathProvider();

            // Base output file name (timestamp will be appended by the provider).
            string outputHtml = "WorkbookExport.html";

            // Export the workbook to HTML using the custom naming pattern.
            workbook.Save(outputHtml, htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
