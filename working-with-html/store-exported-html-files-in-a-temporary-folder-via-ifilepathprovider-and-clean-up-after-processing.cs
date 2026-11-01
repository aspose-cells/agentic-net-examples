// Title: Export a Workbook to HTML in a temporary folder with Aspose.Cells and clean up the files using C#
// Common Searches: how to save Aspose.Cells workbook as HTML to a temporary directory in C# | using a custom IFilePathProvider to obtain a temp folder for Aspose.Cells HTML export | delete temporary folder after exporting workbook to HTML with Aspose.Cells | Aspose.Cells export active worksheet only to HTML5 in a temp folder | retry deleting a temp folder when files are locked after Aspose.Cells HTML save
// Tags: Aspose.Cells export workbook to HTML in temp directory | custom IFilePathProvider for temporary file paths in .NET | HtmlSaveOptions configure HTML5 output and separate image files | cleanup temporary folder after Aspose.Cells HTML export | retry delete temporary directory on IOException C#

using Aspose.Cells;
using System;
using System.IO;
using System.Threading;

// Interface to provide temporary folder paths
public interface IFilePathProvider
{
    // Returns a unique temporary folder path that already exists
    string GetTempFolderPath();
}

// Simple implementation that creates a GUID‑named folder under the system temp directory
// The example defines an IFilePathProvider interface and a TempFilePathProvider that creates a unique GUID‑named folder under the system temp path. HtmlExporter uses this provider to save a Workbook as HTML with HtmlSaveOptions (HTML5, active worksheet only, images as separate files) into the temporary folder, returns the HTML file path, and provides a Cleanup method that safely removes the folder, retrying on IOException or UnauthorizedAccessException.
public class TempFilePathProvider : IFilePathProvider
{
    public string GetTempFolderPath()
    {
        string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempPath);
        return tempPath;
    }
}

// Handles exporting a Workbook to HTML stored in a temporary location
public class HtmlExporter
{
    private readonly IFilePathProvider _pathProvider;

    public HtmlExporter(IFilePathProvider pathProvider)
    {
        _pathProvider = pathProvider;
    }

    // Exports the given workbook to an HTML file inside a temporary folder.
    // Returns the full path to the generated HTML file.
    public string ExportToHtml(Workbook workbook)
    {
        // Obtain a unique temporary folder
        string tempFolder = _pathProvider.GetTempFolderPath();

        // Define the HTML file name inside that folder
        string htmlFilePath = Path.Combine(tempFolder, "output.html");

        // Configure HTML save options
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
        {
            ExportImagesAsBase64 = false,          // Save images as separate files
            ExportActiveWorksheetOnly = true,      // Export only the active sheet
            HtmlVersion = HtmlVersion.Html5        // Use HTML5 output
        };

        // Save the workbook as HTML
        workbook.Save(htmlFilePath, options);

        // Return the path so the caller can use the file and later clean it up
        return htmlFilePath;
    }

    // Deletes the temporary folder that contains the HTML file and any related resources.
    public void Cleanup(string htmlFilePath)
    {
        if (string.IsNullOrEmpty(htmlFilePath))
            return;

        string folder = Path.GetDirectoryName(htmlFilePath);
        if (folder == null || !Directory.Exists(folder))
            return;

        // Attempt to delete the folder; retry after a short pause if files are locked
        try
        {
            Directory.Delete(folder, true);
        }
        catch (IOException)
        {
            Thread.Sleep(500);
            Directory.Delete(folder, true);
        }
        catch (UnauthorizedAccessException)
        {
            Thread.Sleep(500);
            Directory.Delete(folder, true);
        }
    }
}

// Demonstrates usage of the HtmlExporter with temporary storage and cleanup
public class Program
{
    public static void Main()
    {
        // Create or load a workbook
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Hello Aspose.Cells!");

        // Set up the temporary path provider and exporter
        IFilePathProvider pathProvider = new TempFilePathProvider();
        HtmlExporter exporter = new HtmlExporter(pathProvider);

        string htmlPath = null;
        try
        {
            // Export the workbook to HTML stored in a temporary folder
            htmlPath = exporter.ExportToHtml(workbook);

            // Example processing: read and display the HTML content
            string htmlContent = File.ReadAllText(htmlPath);
            Console.WriteLine(htmlContent);
        }
        finally
        {
            // Ensure temporary files are removed regardless of success or failure
            exporter.Cleanup(htmlPath);
        }
    }
}
