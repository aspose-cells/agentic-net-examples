// Title: Generate relative file paths for HTML resources using a custom IFilePathProvider in Aspose.Cells for .NET
// AI Prompts: Create a class that implements IFilePathProvider to store HTML images, CSS, and other resources in a "resources" subdirectory and return the relative path for each file. | Configure HtmlSaveOptions to disable ExportImagesAsBase64, assign the custom file path provider, enable ExportActiveWorksheetOnly, and save the workbook as an HTML file.
// Common Searches: how to export Aspose.Cells workbook to HTML with images saved in a separate folder | using IFilePathProvider to set relative paths for HTML resources in Aspose.Cells .NET | Aspose.Cells HTML export offline browsing resources folder example | disable base64 image embedding when saving Excel as HTML with Aspose.Cells | save only the active worksheet to HTML using Aspose.Cells SaveOptions
// Tags: relative resource folder for Aspose.Cells HTML export | export images as separate files Aspose.Cells HTML | offline HTML browsing with Aspose.Cells workbook | save active worksheet only Aspose.Cells HTMLSaveOptions | custom file path provider for HTML resources Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Custom provider that generates relative paths for HTML resources (images, CSS, etc.)
// The example defines a RelativePathProvider that implements IFilePathProvider to place generated HTML resources (images, CSS, etc.) in a "resources" folder and return their relative paths. HtmlSaveOptions are configured to disable Base64 image embedding, use the custom provider, and export only the active worksheet before saving the workbook as an HTML file.
class RelativePathProvider : IFilePathProvider
{
    // Returns a relative path for the given resource file name.
    public string GetFullName(string fileName)
    {
        const string resourcesFolder = "resources";

        // Ensure the resources folder exists.
        if (!Directory.Exists(resourcesFolder))
        {
            Directory.CreateDirectory(resourcesFolder);
        }

        // Combine folder and file name to create a relative path.
        return Path.Combine(resourcesFolder, fileName);
    }
}

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example content
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells HTML Export with Relative Paths");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export images as separate files (not Base64) so the provider can name them.
                ExportImagesAsBase64 = false,

                // Use the custom file path provider for all generated resources.
                FilePathProvider = new RelativePathProvider(),

                // Export only the active worksheet to keep the output simple.
                ExportActiveWorksheetOnly = true
            };

            // Save the workbook as HTML using the configured options
            workbook.Save("output.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
