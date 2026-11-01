// Title: Save Aspose.Cells workbook as HTML to a UNC network share using a custom IFilePathProvider (C#)
// Common Searches: c# save Aspose.Cells workbook as HTML to a UNC path using IFilePathProvider | how to validate write permissions on a network share before exporting HTML with Aspose.Cells | custom IFilePathProvider implementation for UNC directories in Aspose.Cells | Aspose.Cells HtmlSaveOptions FilePathProvider example for network share | export workbook to HTML with base64 images to a shared folder in C# | Aspose.Cells HTML export to \\server\share folder using custom file path provider
// Tags: custom IFilePathProvider UNC implementation | Aspose.Cells HTML export to network share | validate UNC write access C# | HtmlSaveOptions FilePathProvider usage | export images as base64 Aspose.Cells | save workbook as HTML on shared folder

using System;
using System.IO;
using Aspose.Cells;

// The example implements a NetworkShareFilePathProvider that builds full UNC file paths, ensures directories exist, and verifies write access by creating a temporary file. It then creates a workbook, configures HtmlSaveOptions with ExportImagesAsBase64 and the custom provider, and saves the workbook as HTML directly to a network share (UNC path).
public class NetworkShareFilePathProvider : IFilePathProvider
{
    private readonly string _uncBasePath;

    /// <param name="uncBasePath">The UNC directory where files will be stored.</param>
    public NetworkShareFilePathProvider(string uncBasePath)
    {
        if (string.IsNullOrWhiteSpace(uncBasePath))
            throw new ArgumentException("UNC base path cannot be null or empty.", nameof(uncBasePath));

        // Ensure the path ends without a trailing backslash for consistent concatenation.
        _uncBasePath = uncBasePath.TrimEnd('\\');

        // Validate accessibility of the UNC share.
        ValidateUncPath();
    }

    /// <param name="fileName">File name with extension supplied by Aspose.Cells.</param>
    /// <returns>Full UNC file path.</returns>
    public string GetFullName(string fileName)
    {
        // Combine base UNC path with the supplied file name.
        string fullPath = Path.Combine(_uncBasePath, fileName);

        // Ensure the target directory exists.
        string directory = Path.GetDirectoryName(fullPath);
        if (!Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        return fullPath;
    }

    private void ValidateUncPath()
    {
        // Attempt to create a temporary file to verify write access.
        string testFile = Path.Combine(_uncBasePath, "AsposeCellsAccessTest.tmp");
        try
        {
            // Ensure the directory exists.
            if (!Directory.Exists(_uncBasePath))
                Directory.CreateDirectory(_uncBasePath);

            // Write and delete a temporary file.
            File.WriteAllText(testFile, "test");
            File.Delete(testFile);
        }
        catch (Exception ex)
        {
            throw new IOException($"Unable to access UNC path '{_uncBasePath}'.", ex);
        }
    }
}

class Program
{
    static void Main()
    {
        try
        {
            // Define the UNC network share where HTML files will be saved.
            // Example: \\fileserver\shared\reports
            string uncPath = @"\\fileserver\shared\reports";

            // Ensure the UNC directory exists before proceeding.
            if (!Directory.Exists(uncPath))
                Directory.CreateDirectory(uncPath);

            // Instantiate the custom file path provider.
            var filePathProvider = new NetworkShareFilePathProvider(uncPath);

            // Create a new workbook and populate it with sample data.
            var workbook = new Workbook();
            var cells = workbook.Worksheets[0].Cells;
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Quantity");
            cells["A2"].PutValue("Apples");
            cells["B2"].PutValue(120);
            cells["A3"].PutValue("Bananas");
            cells["B3"].PutValue(85);

            // Configure HTML save options to use the custom file path provider.
            var htmlOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true,
                // The provider will be called for each auxiliary file (e.g., images, CSS).
                FilePathProvider = filePathProvider,
                // Embed images as base64 to avoid separate image files.
                ExportImagesAsBase64 = true
                // Note: ExportCssSeparately is not available in this version of Aspose.Cells.
            };

            // Define the main HTML file name (including extension).
            string htmlFileName = "SalesReport.html";
            string htmlFullPath = Path.Combine(uncPath, htmlFileName);

            // Save the workbook as HTML to the UNC location.
            try
            {
                workbook.Save(htmlFullPath, htmlOptions);
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error during saving HTML: {saveEx.Message}");
                return;
            }

            // Verify that the main HTML file was created successfully.
            if (File.Exists(htmlFullPath))
            {
                Console.WriteLine($"HTML file successfully saved to UNC path: {htmlFullPath}");
            }
            else
            {
                Console.WriteLine("Failed to save the HTML file to the UNC location.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
