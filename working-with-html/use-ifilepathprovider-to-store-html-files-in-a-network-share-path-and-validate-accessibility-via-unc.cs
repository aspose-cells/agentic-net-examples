// Title: Save Aspose.Cells Worksheets as HTML to a UNC Share using a Custom IFilePathProvider (C#)
// Description: Demonstrates how to implement IFilePathProvider to generate a UNC path for each worksheet HTML file, verify share accessibility, and configure HtmlSaveOptions so the main workbook HTML is saved locally while individual sheet files are stored on a network share.
// Keywords: Aspose.Cells IFilePathProvider | UNC network share | save worksheet html | HtmlSaveOptions custom provider | C# network share validation | store Excel HTML on server | Aspose.Cells HTML export
// Common Searches: Aspose.Cells custom IFilePathProvider example | save each worksheet as separate HTML files | export Aspose.Cells HTML to UNC path | validate network share before saving Excel HTML | C# HtmlSaveOptions file path provider UNC
// Developer Intent: Export each worksheet to its own HTML file on a UNC share and confirm the share is reachable before writing.
// Use Cases: Create per‑sheet HTML reports on a shared server for team collaboration. | Automate nightly export of workbook worksheets to a central UNC folder for downstream processing. | Pre‑check write permissions on a network share to prevent runtime failures during large exports.
// AI Prompts: Show a C# implementation of IFilePathProvider that writes Aspose.Cells worksheet HTML files to a UNC location and includes robust error handling. | Explain how to add detailed logging for UNC folder accessibility checks in the custom file path provider. | Demonstrate configuring HtmlSaveOptions to use a custom IFilePathProvider while specifying a different output folder for the main HTML file.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to implement IFilePathProvider to generate a UNC path for each worksheet HTML file, verify share accessibility, and configure HtmlSaveOptions so the main workbook HTML is saved locally while individual sheet files are stored on a network share.
class NetworkShareFilePathProvider : IFilePathProvider
{
    private readonly string _uncFolder;

    public NetworkShareFilePathProvider(string uncFolder)
    {
        _uncFolder = uncFolder;
    }

    // Returns the full UNC path for each worksheet HTML file
    public string GetFullName(string sheetName)
    {
        string fileName = $"{sheetName}.html";
        return Path.Combine(_uncFolder, fileName);
    }

    // Simple validation that the UNC folder is reachable and writable
    public bool IsAccessible()
    {
        try
        {
            if (!Directory.Exists(_uncFolder))
                return false;

            string testFile = Path.Combine(_uncFolder, "access_test.tmp");
            using (FileStream fs = File.Create(testFile)) { }
            File.Delete(testFile);
            return true;
        }
        catch
        {
            return false;
        }
    }
}

class Program
{
    static void Main()
    {
        // UNC path to the network share where individual sheet HTML files will be stored
        string uncPath = @"\\MyServer\SharedFolder\ExcelHtml";

        // Initialize the custom file path provider
        var provider = new NetworkShareFilePathProvider(uncPath);

        // Validate that the UNC location is accessible before proceeding
        if (!provider.IsAccessible())
        {
            Console.WriteLine($"Unable to access network share: {uncPath}");
            return;
        }

        // Create a workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Report";
        sheet.Cells["A1"].PutValue("Hello from UNC share");

        // Configure HTML save options to use the custom provider
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.FilePathProvider = provider;
        // Export all worksheets (each will be saved to the UNC folder)
        saveOptions.ExportActiveWorksheetOnly = false;

        // Save the workbook; the main HTML file is created locally,
        // while each worksheet's HTML is stored in the UNC location
        workbook.Save("MainReport.html", saveOptions);

        Console.WriteLine("Workbook saved. Individual worksheet HTML files are stored in the UNC share.");
    }
}
