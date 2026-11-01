// Title: Use a custom IFilePathProvider with HtmlSaveOptions to repair worksheet hyperlink URLs during Excel-to-HTML export in Aspose.Cells for .NET
// AI Prompts: Implement a class that inherits IFilePathProvider, override GetFullName to prepend a virtual folder (e.g., "files/"), and assign this instance to HtmlSaveOptions.FilePathProvider before saving the workbook as HTML. | Update the HTML export routine to supply a custom file‑path provider so that all worksheet hyperlinks in the generated HTML point to the desired relative location.
// Common Searches: Aspose.Cells how to change hyperlink URLs in HTML export using IFilePathProvider | C# set HtmlSaveOptions.FilePathProvider to fix broken links in exported HTML | custom file path provider for Excel to HTML conversion Aspose.Cells .NET | adjust relative hyperlink paths when saving workbook as HTML with Aspose.Cells | example of implementing IFilePathProvider for HTML export in Aspose.Cells
// Tags: HtmlSaveOptions FilePathProvider usage | fix broken hyperlinks in HTML export | Aspose.Cells hyperlink path customization | Excel to HTML conversion hyperlink handling | custom IFilePathProvider implementation

using System;
using System.IO;
using Aspose.Cells;

// Custom implementation of IFilePathProvider to control hyperlink file paths during HTML export
// Shows how to create a CustomFilePathProvider that prefixes hyperlink file names with a virtual folder, assign it to HtmlSaveOptions.FilePathProvider, and save a workbook as HTML so that worksheet hyperlinks reference the corrected paths.
public class CustomFilePathProvider : IFilePathProvider
{
    // Returns the path that should be used in the generated HTML for a given file name.
    public string GetFullName(string fileName)
    {
        // Example logic: prepend a virtual folder or modify the path as needed.
        return $"files/{fileName}";
    }
}

public class HtmlExportWithCustomFilePathProvider
{
    public static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook.
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Assign the custom IFilePathProvider to fix broken worksheet hyperlinks.
                FilePathProvider = new CustomFilePathProvider()
            };

            const string outputPath = "output.html";

            // Save the workbook as HTML using the configured options.
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
