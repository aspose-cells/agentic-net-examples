// Title: Save an Excel workbook as HTML with timestamped file and auxiliary names using a custom IFilePathProvider in Aspose.Cells for .NET
// Common Searches: how to use IFilePathProvider to add timestamps to HTML export files in Aspose.Cells C# | Aspose.Cells export workbook to HTML with custom file naming for images and CSS | C# generate timestamped HTML file name when saving Excel as HTML with Aspose.Cells | customize auxiliary file names during Aspose.Cells HTML conversion using IFilePathProvider | timestamped output file pattern for Aspose.Cells HtmlSaveOptions in .NET
// Tags: Aspose.Cells IFilePathProvider implementation | HTML export with timestamped auxiliary resources | C# HtmlSaveOptions custom file path provider | Excel workbook to HTML with unique file names | Aspose.Cells timestamped output naming

using System;
using System.IO;
using Aspose.Cells;

// Custom file path provider that appends a timestamp to each generated file name (e.g., images, CSS)
// The example implements a TimestampFilePathProvider that appends a yyyyMMdd_HHmmss timestamp to every generated file name (HTML, images, CSS) during HTML conversion. HtmlSaveOptions is configured with this provider, a timestamped main HTML file name is created, and the workbook is saved as HTML, handling missing source files gracefully.
class TimestampFilePathProvider : IFilePathProvider
{
    // Returns a relative path (used by Aspose.Cells when generating auxiliary files)
    public string GetFilePath(string fileName)
    {
        // Build a timestamp string
        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

        // Separate directory, name and extension
        string directory = Path.GetDirectoryName(fileName);
        string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
        string extension = Path.GetExtension(fileName);

        // Create new file name with timestamp
        string newFileName = $"{nameWithoutExt}_{timestamp}{extension}";

        // Preserve original directory if any
        return Path.Combine(string.IsNullOrEmpty(directory) ? "" : directory, newFileName);
    }

    // Returns the full physical path for the generated file
    public string GetFullName(string fileName)
    {
        // Resolve the relative path returned by GetFilePath to an absolute path
        string relativePath = GetFilePath(fileName);
        return Path.GetFullPath(relativePath);
    }
}

class Program
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Ensure the input file exists; otherwise create an empty workbook
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file '{inputPath}' not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Configure HTML save options with the custom provider
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                FilePathProvider = new TimestampFilePathProvider()
            };

            // Create a timestamped name for the main HTML file
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string outputHtmlPath = $"output_{timestamp}.html";

            // Export the workbook to HTML
            workbook.Save(outputHtmlPath, htmlOptions);
            Console.WriteLine($"Workbook saved as HTML to '{outputHtmlPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
