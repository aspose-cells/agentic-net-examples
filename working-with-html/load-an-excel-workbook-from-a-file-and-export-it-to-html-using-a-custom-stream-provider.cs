// Title: Save an Excel workbook as HTML with a custom IStreamProvider using Aspose.Cells for .NET
// AI Prompts: Create a C# program that loads an .xlsx file (or creates a new workbook), implements an IStreamProvider to route HTML resources to a folder, and saves the workbook to a MemoryStream via Aspose.Cells. | Write C# code that extracts the generated HTML bytes from the MemoryStream, ensures the target directory exists, and writes the bytes to a .html file on disk.
// Common Searches: how to export Excel to HTML with a custom resource folder using Aspose.Cells C# | Aspose.Cells HtmlSaveOptions StreamProvider example .NET | save workbook as HTML to memory stream then to file Aspose.Cells | custom IStreamProvider implementation for HTML resources Aspose.Cells | C# generate HTML from Excel and store images in separate folder Aspose.Cells
// Tags: Aspose.Cells HTML export with custom stream handling | C# convert generated HTML bytes to file | export Excel workbook to HTML while preserving images | configure HTML export settings for resource folder | using stream provider to write HTML assets

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing .xlsx file (or creates a new workbook), configures HtmlSaveOptions.StreamProvider with a custom IStreamProvider that writes HTML resources to a designated folder, saves the workbook to a MemoryStream, and writes the resulting HTML bytes to output.html.
class CustomStreamProvider : IStreamProvider
{
    private readonly string _resourceFolder;

    public CustomStreamProvider(string resourceFolder)
    {
        _resourceFolder = resourceFolder;
        if (!Directory.Exists(_resourceFolder))
            Directory.CreateDirectory(_resourceFolder);
    }

    // Called before any stream operations start
    public void InitStream(StreamProviderOptions options)
    {
        // No initialization required for this simple implementation
    }

    // Provides a stream for a given resource name
    public Stream GetStream(string name, StreamProviderOptions options)
    {
        string filePath = Path.Combine(_resourceFolder, name);
        return new FileStream(filePath, FileMode.Create, FileAccess.Write);
    }

    // Called after all stream operations are finished
    public void CloseStream(StreamProviderOptions options)
    {
        // No cleanup required for this simple implementation
    }
}

class Program
{
    static void Main()
    {
        try
        {
            // Load the Excel workbook from a file if it exists; otherwise create a new workbook
            string inputFile = "input.xlsx";
            Workbook workbook = File.Exists(inputFile) ? new Workbook(inputFile) : new Workbook();

            // Configure HTML save options to use the custom stream provider
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.StreamProvider = new CustomStreamProvider("HtmlResources");

            // Save the workbook as HTML into a memory stream
            using (MemoryStream htmlStream = new MemoryStream())
            {
                workbook.Save(htmlStream, htmlOptions);

                // Ensure the output directory exists
                string outputPath = "output.html";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);

                // Write the generated HTML to a physical file
                File.WriteAllBytes(outputPath, htmlStream.ToArray());
            }

            Console.WriteLine("HTML export completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
