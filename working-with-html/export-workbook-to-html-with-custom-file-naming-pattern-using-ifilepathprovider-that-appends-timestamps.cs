// Title: Save an Excel workbook as HTML with a timestamped filename using Aspose.Cells IFilePathProvider in C#
// Common Searches: how to generate timestamped HTML file names when saving Excel with Aspose.Cells C# | custom IFilePathProvider example for Aspose.Cells HTML export | Aspose.Cells save workbook to HTML with dynamic file name pattern | C# Aspose.Cells HtmlSaveOptions file path provider timestamp | append current datetime to exported HTML file name using Aspose.Cells
// Tags: Aspose.Cells IFilePathProvider timestamp | HTML export filename customization Aspose.Cells | HtmlSaveOptions custom file path provider | C# dynamic HTML output name Aspose.Cells | Excel to HTML conversion with timestamped output

using System;
using System.IO;
using Aspose.Cells;

// The sample implements a TimestampFilePathProvider that appends a yyyyMMddHHmmssfff timestamp to the base name of the HTML file. It loads an Excel workbook, configures HtmlSaveOptions with this provider, and saves the workbook as HTML, producing a uniquely timestamped HTML file.
class TimestampFilePathProvider : IFilePathProvider
{
    // Generates a file path with a timestamp appended to the base name.
    public string GetFilePath(string fileName, string fileExtension, int index)
    {
        string timestamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");
        string directory = Path.GetDirectoryName(fileName);
        string baseName = Path.GetFileNameWithoutExtension(fileName);
        string newFileName = $"{baseName}_{timestamp}{fileExtension}";
        return Path.Combine(string.IsNullOrEmpty(directory) ? "" : directory, newFileName);
    }

    // Required by IFilePathProvider – returns the original file name unchanged.
    public string GetFullName(string fileName)
    {
        return fileName;
    }
}

class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options and assign the custom file path provider.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            htmlOptions.FilePathProvider = new TimestampFilePathProvider();

            // Export the workbook to HTML using the custom naming pattern.
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
