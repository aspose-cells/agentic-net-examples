// Title: Generate relative resource paths for offline HTML export with IFilePathProvider in Aspose.Cells for .NET
// Common Searches: how to implement IFilePathProvider to set relative paths for HTML resources in Aspose.Cells C# | Aspose.Cells HtmlSaveOptions custom file path provider example for offline HTML | save Excel workbook as HTML with a resources subfolder using Aspose.Cells .NET | C# convert workbook to HTML and place images/scripts in a relative folder | using IFilePathProvider to control external file locations during HTML export in Aspose.Cells
// Tags: custom IFilePathProvider for HTML resources | Aspose.Cells HtmlSaveOptions relative path provider | offline HTML export with resource folder | C# Excel to HTML conversion with external assets | relative path generation for HTML assets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Custom provider that generates relative paths for HTML resources
// The example defines a RelativePathProvider that implements IFilePathProvider, returning a "resources/filename" relative path and its full physical path. It creates a sample workbook if needed, ensures a "resources" directory exists, configures HtmlSaveOptions to use the custom provider, and saves the workbook as an HTML file with all external resources placed in the relative folder for offline browsing.
class RelativePathProvider : IFilePathProvider
{
    // Returns a relative path (e.g., "resources/filename") for each external file
    public string GetFilePath(string fileName)
    {
        return Path.Combine("resources", fileName);
    }

    // Returns the full physical path for the given file name
    public string GetFullName(string fileName)
    {
        return Path.GetFullPath(GetFilePath(fileName));
    }
}

class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputHtml = "output.html";

        try
        {
            // Ensure the input workbook exists; create a simple one if missing
            if (!File.Exists(inputPath))
            {
                var wb = new Workbook();
                wb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                wb.Save(inputPath);
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure the resources directory exists for external files
            string resourcesDir = Path.Combine(Directory.GetCurrentDirectory(), "resources");
            if (!Directory.Exists(resourcesDir))
                Directory.CreateDirectory(resourcesDir);

            // Configure HTML save options to use the custom file path provider
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            htmlOptions.FilePathProvider = new RelativePathProvider();

            // Save the workbook as an HTML file
            workbook.Save(outputHtml, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputHtml}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
