// Title: Create relative resource paths for HTML export with IFilePathProvider in Aspose.Cells for .NET
// Common Searches: how to use IFilePathProvider to generate relative image links when saving workbook to HTML in Aspose.Cells C# | Aspose.Cells C# export workbook to HTML with resources in a separate folder for offline viewing | set HtmlSaveOptions.IsFullPathLink false and provide custom file path provider in Aspose.Cells | C# example of creating Resources folder for HTML assets using Aspose.Cells IFilePathProvider | Aspose.Cells generate relative paths for CSS and images during HTML conversion
// Tags: IFilePathProvider custom implementation for HTML resources | Aspose.Cells HtmlSaveOptions relative links | export workbook to HTML with Resources folder | C# generate offline‑friendly HTML assets Aspose.Cells | relative image paths in Aspose.Cells HTML export

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Custom file path provider that generates relative paths for HTML resources.
    // Demonstrates a custom RelativePathProvider that places HTML resources (images, CSS, etc.) in a "Resources" subfolder, returns relative URLs for those files, ensures the folder exists, and configures HtmlSaveOptions with IsFullPathLink set to false to produce offline‑compatible HTML output.
    public class RelativePathProvider : IFilePathProvider
    {
        // Called for each resource (image, css, etc.) to get a relative path used in HTML.
        public string GetFilePath(string fileName, string fileExtension, int index)
        {
            // Place all resources in a subfolder named "Resources" next to the HTML file.
            return $"Resources/{fileName}{fileExtension}";
        }

        // Returns the full physical path where the resource file will be written.
        public string GetFullName(string fileName)
        {
            // Ensure the Resources folder exists.
            string resourcesDir = Path.Combine(Directory.GetCurrentDirectory(), "Resources");
            if (!Directory.Exists(resourcesDir))
            {
                Directory.CreateDirectory(resourcesDir);
            }

            // Combine the folder with the file name (fileName already contains extension).
            return Path.Combine(resourcesDir, fileName);
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "InputWorkbook.xlsx";
                Workbook workbook;

                // Load existing workbook if present; otherwise create a simple one.
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    workbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                }

                // Configure HTML save options.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportImagesAsBase64 = false, // Generate separate image files.
                    IsFullPathLink = false,       // Use relative links.
                    FilePathProvider = new RelativePathProvider()
                };

                // Save the workbook as HTML.
                const string outputHtml = "OutputWorkbook.html";
                workbook.Save(outputHtml, saveOptions);
                Console.WriteLine($"Workbook successfully saved to '{outputHtml}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
