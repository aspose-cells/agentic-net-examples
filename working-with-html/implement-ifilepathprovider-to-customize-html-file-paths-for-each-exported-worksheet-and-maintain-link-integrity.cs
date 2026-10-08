// Title: How to use IFilePathProvider to save each worksheet as an HTML file in a subfolder with Aspose.Cells for .NET
// AI Prompts: Implement a class that inherits IFilePathProvider and returns a subdirectory path for every worksheet’s HTML file during export. | Configure HtmlSaveOptions.FilePathProvider with the custom provider and export a workbook so that index.html correctly links to the sheet files stored in the subfolder. | Add a pre‑export verification that the source Excel file exists and wrap the export call in try‑catch for graceful error handling.
// Common Searches: Aspose.Cells .NET custom IFilePathProvider example for HTML export | save each worksheet as separate HTML file in a specific folder using Aspose.Cells | how to keep index.html links working when exporting Excel to multiple HTML files | HtmlSaveOptions FilePathProvider usage for organizing exported sheets | export Excel workbook to HTML with subfolder for sheet files in C#
// Tags: IFilePathProvider custom path | separate HTML files per worksheet | HtmlSaveOptions file path customization | organize exported HTML sheets folder | maintain sheet link integrity HTML

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// Custom implementation of IFilePathProvider to control HTML file paths per worksheet
// The example defines a CustomFilePathProvider that implements IFilePathProvider to place each worksheet's HTML output into an "ExportedSheets" subfolder. HtmlSaveOptions.FilePathProvider is set to this custom class, enabling separate HTML files per sheet while preserving the index page links. The code also checks for the existence of the source workbook and includes basic exception handling.
public class CustomFilePathProvider : IFilePathProvider
{
    // This method is called for each worksheet during HTML export.
    // It receives the default file name (e.g., "Sheet1.html") and returns the full path where the file will be saved.
    public string GetFullName(string fileName)
    {
        // Create a subfolder to keep exported sheets organized.
        string folderPath = Path.Combine("ExportedSheets");
        Directory.CreateDirectory(folderPath);

        // Return the combined path.
        return Path.Combine(folderPath, fileName);
    }
}

public class HtmlExportExample
{
    public static void Main()
    {
        const string inputPath = "input.xlsx";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook.
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export each worksheet to a separate HTML file.
                // The default behavior when ExportToSingleFile is not set is to create separate files.
                // Assign the custom file path provider.
                FilePathProvider = new CustomFilePathProvider()
            };

            // Save the workbook to HTML using the configured options.
            // The first argument is the main entry HTML file (index) that contains links to individual sheets.
            workbook.Save("index.html", saveOptions);

            Console.WriteLine("Workbook successfully exported to HTML.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
