// Title: Export Aspose.Cells Workbook to HTML in a Temporary Folder with IFilePathProvider (C#)
// Description: Shows how to create a unique temporary directory, set HtmlSaveOptions with a custom IFilePathProvider so the main HTML and worksheet files are saved there, export the workbook, and then safely delete the temporary folder.
// Keywords: Aspose.Cells | HtmlSaveOptions | IFilePathProvider | temporary directory | C# | .NET | export workbook to HTML | custom file path provider | clean up temporary files | delete temp folder
// Common Searches: Aspose.Cells export HTML to specific folder C# | IFilePathProvider example Aspose.Cells | how to delete temporary folder after HTML export Aspose.Cells | save workbook as HTML in temp directory .NET | Aspose.Cells temporary folder cleanup
// Developer Intent: Create a temp folder, export a workbook to HTML using a custom IFilePathProvider, and remove the folder afterward.
// Use Cases: Generate an HTML preview of a spreadsheet for email or API response without leaving permanent files on disk. | Render workbook content in a web app, store HTML files in a temporary location for processing, then clean up after the request completes. | Isolate HTML export artifacts in automated unit or integration tests that require a clean file system state. | Batch‑process multiple workbooks to HTML on a server, using per‑run temporary folders to avoid naming conflicts.
// AI Prompts: Write C# code that creates a temporary directory, uses Aspose.Cells HtmlSaveOptions with a custom IFilePathProvider to export a workbook to HTML in that directory, and then deletes the directory. | Explain step‑by‑step how to implement IFilePathProvider to direct worksheet‑specific HTML files to a given folder and ensure proper cleanup with Directory.Delete. | Provide a concise guide for exporting a workbook to HTML in a temporary location and safely removing all generated files in a .NET application.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to create a unique temporary directory, set HtmlSaveOptions with a custom IFilePathProvider so the main HTML and worksheet files are saved there, export the workbook, and then safely delete the temporary folder.
class Program
{
    static void Main()
    {
        // Create a unique temporary directory for the exported HTML files
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeHtmlExport_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempFolder);

        // Create a workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Hello");
        worksheet.Cells["A2"].PutValue("World");

        // Configure HtmlSaveOptions to use a custom IFilePathProvider that writes to the temp folder
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.FilePathProvider = new TempFolderFilePathProvider(tempFolder);

        // Save the workbook as HTML; the main file and any worksheet‑specific files will be placed in the temp folder
        string mainHtmlPath = Path.Combine(tempFolder, "index.html");
        workbook.Save(mainHtmlPath, saveOptions);

        Console.WriteLine($"HTML exported to temporary folder: {tempFolder}");

        // Perform any additional processing here if needed...

        // Clean up the temporary directory after processing
        Directory.Delete(tempFolder, true);
        Console.WriteLine("Temporary folder cleaned up.");
    }

    // Custom implementation of IFilePathProvider that directs worksheet HTML files to the specified temporary folder
    private class TempFolderFilePathProvider : IFilePathProvider
    {
        private readonly string _folder;

        public TempFolderFilePathProvider(string folder)
        {
            _folder = folder;
        }

        public string GetFullName(string sheetName)
        {
            // Generate a full path for each worksheet HTML file inside the temporary folder
            return Path.Combine(_folder, $"{sheetName}.html");
        }
    }
}
