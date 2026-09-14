// Title: Export an Excel workbook to HTML using Aspose.Cells, store the file in a temporary folder via IFilePathProvider, and clean up the folder afterward
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, saves it as an HTML file in a temporary directory obtained from an IFilePathProvider implementation, and returns the generated HTML path. | Add a method that removes the temporary directory and all its contents after the HTML file has been processed. | Show how to inject a mock IFilePathProvider to unit‑test the HTML export and cleanup logic.
// Common Searches: aspnet core export excel to html using aspose.cells and temporary directory | c# how to delete temporary directory after aspose.cells html conversion | implement IFilePathProvider for temporary file storage with Aspose.Cells | unit test Aspose.Cells HTML export with mock IFilePathProvider
// Tags: Aspose.Cells HTML export to temp location | IFilePathProvider custom temp directory | C# cleanup temp folder after Aspose.Cells save | unit testing Aspose.Cells HTML conversion | manage transient files with Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Interface to provide temporary folder paths
public interface IFilePathProvider
{
    string GetTemporaryFolder();
}

// Simple implementation that creates a unique temp folder
// // Demonstrates an IFilePathProvider that creates a unique temporary folder, uses Aspose.Cells to load an Excel workbook and save it as HTML inside that folder, returns the HTML file path, and provides a CleanUp method to delete the temporary folder and its contents after processing.
public class TempFolderProvider : IFilePathProvider
{
    public string GetTemporaryFolder()
    {
        string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempPath);
        return tempPath;
    }
}

// Handles exporting a workbook to HTML and cleaning up the temp folder
public class HtmlExporter
{
    private readonly IFilePathProvider _pathProvider;

    public HtmlExporter(IFilePathProvider pathProvider)
    {
        _pathProvider = pathProvider;
    }

    // Exports the workbook located at workbookPath to an HTML file in a temporary folder
    // Returns the full path of the generated HTML file
    public string ExportWorkbookToHtml(string workbookPath)
    {
        // Load the workbook using Aspose.Cells
        Workbook workbook = new Workbook(workbookPath);

        // Obtain a temporary folder from the provider
        string tempFolder = _pathProvider.GetTemporaryFolder();

        // Define the output HTML file name and full path
        string htmlFilePath = Path.Combine(tempFolder, "exported.html");

        // Save the workbook as HTML
        workbook.Save(htmlFilePath, SaveFormat.Html);

        // Return the path for further processing
        return htmlFilePath;
    }

    // Deletes the temporary folder and all its contents
    public void CleanUp(string folderPath)
    {
        if (Directory.Exists(folderPath))
        {
            Directory.Delete(folderPath, true);
        }
    }
}

// Example usage
class Program
{
    static void Main()
    {
        // Initialize the provider and exporter
        IFilePathProvider provider = new TempFolderProvider();
        HtmlExporter exporter = new HtmlExporter(provider);

        // Path to the source Excel file
        string sourceExcel = @"C:\Data\sample.xlsx";

        // Export the workbook to HTML in a temporary folder
        string htmlPath = exporter.ExportWorkbookToHtml(sourceExcel);

        // Use the HTML file as needed (e.g., display, send, etc.)
        Console.WriteLine("HTML exported to: " + htmlPath);

        // After processing, clean up the temporary folder
        string tempFolder = Path.GetDirectoryName(htmlPath);
        exporter.CleanUp(tempFolder);
    }
}
