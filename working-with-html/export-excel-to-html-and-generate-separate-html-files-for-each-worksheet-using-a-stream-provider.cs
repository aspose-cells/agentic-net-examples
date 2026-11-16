// Title: Export each worksheet of an Excel workbook to individual HTML files with Aspose.Cells using a custom IStreamProvider (C#)
// Common Searches: C# Aspose.Cells export single worksheet to HTML using ExportActiveWorksheetOnly | How to save Excel worksheets as separate HTML files with resource folders in Aspose.Cells | Implement custom IStreamProvider for HTML resources when converting Excel to HTML in .NET | Aspose.Cells generate per‑sheet HTML output with images and CSS saved to folder | MemoryStream based HTML export for each sheet using Aspose.Cells C# example
// Tags: ExportActiveWorksheetOnly HtmlSaveOptions | custom IStreamProvider Aspose.Cells | per‑sheet HTML export C# | Excel to HTML resource folder Aspose | worksheet HTML generation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an Excel workbook, iterates through its worksheets, activates each one, and uses HtmlSaveOptions with ExportActiveWorksheetOnly set to true together with a custom FileStreamProvider that creates a dedicated resources folder for the sheet. Each worksheet is saved to a MemoryStream and written to a .html file named after the sheet, while embedded resources such as images and CSS are stored in the corresponding folder.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputFile);

            // Iterate through each worksheet and export it as a separate HTML file
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Activate the current worksheet so ExportActiveWorksheetOnly works correctly
                workbook.Worksheets.ActiveSheetIndex = sheet.Index;

                // Configure HTML save options with a custom stream provider for resources
                HtmlSaveOptions options = new HtmlSaveOptions
                {
                    ExportActiveWorksheetOnly = true,
                    StreamProvider = new FileStreamProvider(sheet.Name)
                };

                // Save the worksheet HTML into a memory stream
                using (MemoryStream htmlStream = new MemoryStream())
                {
                    workbook.Save(htmlStream, options);
                    htmlStream.Position = 0; // Reset stream position for reading

                    // Write the HTML content to a file named after the worksheet
                    string htmlFileName = $"{sheet.Name}.html";
                    File.WriteAllText(htmlFileName, new StreamReader(htmlStream).ReadToEnd());
                }
            }

            Console.WriteLine("HTML files generated successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Custom stream provider that writes embedded resources to a folder specific to each worksheet
class FileStreamProvider : IStreamProvider
{
    private readonly string _resourceFolder;

    public FileStreamProvider(string sheetName)
    {
        // Create a folder to store resources for this worksheet (e.g., images, CSS)
        _resourceFolder = $"{sheetName}_resources";
        Directory.CreateDirectory(_resourceFolder);
    }

    // Called by Aspose.Cells when it needs a stream for a resource file
    public Stream GetStream(string fileName, StreamProviderOptions options)
    {
        string fullPath = Path.Combine(_resourceFolder, fileName);
        return new FileStream(fullPath, FileMode.Create, FileAccess.Write);
    }

    // Called after the resource has been written; no additional cleanup needed here
    public void CloseStream(string fileName, StreamProviderOptions options)
    {
        // No action required
    }

    // Optional initialization method; not needed for this implementation
    public void InitStream(StreamProviderOptions options)
    {
        // No action required
    }

    // Optional cleanup method; not needed for this implementation
    public void CloseStream(StreamProviderOptions options)
    {
        // No action required
    }
}
