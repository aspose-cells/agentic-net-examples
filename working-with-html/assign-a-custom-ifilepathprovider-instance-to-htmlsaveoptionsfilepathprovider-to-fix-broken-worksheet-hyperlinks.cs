// Title: Use a custom IFilePathProvider with HtmlSaveOptions to keep worksheet hyperlinks intact when exporting Excel to HTML in Aspose.Cells for .NET
// AI Prompts: Create a class that implements IFilePathProvider returning the original file name, assign it to HtmlSaveOptions.FilePathProvider, and export a Workbook to HTML so that hyperlinks are preserved. | Modify an existing HTML export routine to inject a custom file path provider, ensuring linked resources resolve correctly in the generated HTML file.
// Common Searches: Aspose.Cells how to maintain hyperlink URLs when saving workbook as HTML | C# set HtmlSaveOptions.FilePathProvider to custom implementation for Excel to HTML conversion | example of IFilePathProvider for preserving links in Aspose.Cells HTML export | fix broken worksheet hyperlinks after exporting Excel to HTML with Aspose.Cells .NET
// Tags: custom IFilePathProvider implementation Aspose.Cells | HtmlSaveOptions FilePathProvider C# | preserve hyperlinks during Excel to HTML export | Aspose.Cells HTML export hyperlink handling | C# Excel workbook to HTML with custom file path provider

using Aspose.Cells;
using System;
using System.IO;

// Shows how to implement a simple IFilePathProvider that returns the file name unchanged, assign it to HtmlSaveOptions.FilePathProvider, and save an Excel workbook as HTML while keeping worksheet hyperlinks functional.
public class CustomFilePathProvider : IFilePathProvider
{
    // Returns the full path for a given file name during HTML export.
    // This simple implementation returns the file name unchanged.
    public string GetFullName(string fileName)
    {
        return fileName;
    }
}

public class HtmlExportWithCustomFilePathProvider
{
    public static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        try
        {
            // Verify that the input workbook exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Assign the custom file path provider to handle hyperlinks.
                FilePathProvider = new CustomFilePathProvider()
            };

            // Save the workbook as HTML.
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
