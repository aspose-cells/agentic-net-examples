// Title: Create a custom IFilePathProvider in Aspose.Cells for .NET to generate lowercase, hyphen‑separated worksheet URLs when saving to HTML
// AI Prompts: Write a C# class that implements Aspose.Cells.IFilePathProvider, converting worksheet names to lowercase hyphen‑separated strings for HTML file names. | Show how to set HtmlSaveOptions.FilePathProvider to the custom provider and export a workbook to HTML. | Add null‑check handling in GetFilePath so empty worksheet names are returned unchanged.
// Common Searches: how to use IFilePathProvider in Aspose.Cells to customize HTML file names | Aspose.Cells export Excel to HTML with lowercase worksheet URLs | C# convert Excel sheet name to URL friendly format using Aspose.Cells | custom file path provider example for Aspose.Cells HTMLSaveOptions
// Tags: Aspose.Cells custom IFilePathProvider for HTML export | lowercase hyphenated worksheet file names | HTMLSaveOptions file path customization | C# URL‑friendly Excel sheet names | worksheet name to URL conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Implements IFilePathProvider to convert worksheet names to lowercase URLs.
// Implements IFilePathProvider to transform worksheet names into lowercase, hyphen‑separated strings, enabling consistent, URL‑friendly file names when saving a workbook to HTML with Aspose.Cells.
public class LowercaseWorksheetPathProvider : IFilePathProvider
{
    // Returns a lowercase, URL‑friendly version of the file name.
    public string GetFilePath(string fileName)
    {
        if (string.IsNullOrEmpty(fileName))
            return fileName;

        string lower = fileName.ToLowerInvariant();
        lower = lower.Replace(' ', '-');
        return lower;
    }

    // Returns the full path for the given file name.
    // For this example we simply reuse GetFilePath; adjust as needed for real scenarios.
    public string GetFullName(string fileName)
    {
        return GetFilePath(fileName);
    }
}

// Example usage.
public class Example
{
    public static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string outputPath = "Output.html";

        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load an existing workbook.
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options with the custom file path provider.
            var pathProvider = new LowercaseWorksheetPathProvider();
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                FilePathProvider = pathProvider
            };

            // Demonstrate retrieving the path for the first worksheet using the provider.
            string worksheetName = workbook.Worksheets[0].Name;
            string path = pathProvider.GetFilePath(worksheetName);
            Console.WriteLine($"Lowercase URL for worksheet '{worksheetName}': {path}");

            // Save the workbook as HTML (the provider affects generated file names/paths).
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
