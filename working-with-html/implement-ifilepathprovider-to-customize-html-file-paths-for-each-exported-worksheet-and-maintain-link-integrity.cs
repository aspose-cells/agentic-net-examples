using System;
using System.IO;
using Aspose.Cells;

// Custom provider that generates a full file path for each worksheet HTML file
public class CustomFilePathProvider : IFilePathProvider
{
    private readonly string _baseFolder;

    public CustomFilePathProvider(string baseFolder)
    {
        _baseFolder = baseFolder;
        // Ensure the output directory exists
        Directory.CreateDirectory(_baseFolder);
    }

    // Returns the absolute path for the given worksheet name
    public string GetFullName(string sheetName)
    {
        // Example: C:\MyProject\HtmlSheets\Sheet1.html
        string fileName = $"{sheetName}.html";
        return Path.Combine(_baseFolder, fileName);
    }
}

public class ExportWorkbookToHtmlWithCustomPaths
{
    public static void Run()
    {
        // Create a workbook with two worksheets and some sample data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Name = "SheetOne";
        workbook.Worksheets[0].Cells["A1"].PutValue("Hello from Sheet One");

        workbook.Worksheets.Add("SheetTwo");
        workbook.Worksheets[1].Cells["A1"].PutValue("Hello from Sheet Two");

        // Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // Use full path links so that inter‑sheet references remain valid
            IsFullPathLink = true,
            // Set the custom file path provider
            FilePathProvider = new CustomFilePathProvider(Path.Combine(Environment.CurrentDirectory, "HtmlSheets"))
        };

        // Ensure the output folder exists (provider already creates it, but double‑check)
        string outputFolder = Path.Combine(Environment.CurrentDirectory, "HtmlSheets");
        Directory.CreateDirectory(outputFolder);

        // Save the workbook; the main HTML file will contain links to the separate sheet files
        string mainHtmlPath = Path.Combine(outputFolder, "Workbook.html");
        workbook.Save(mainHtmlPath, saveOptions);
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            ExportWorkbookToHtmlWithCustomPaths.Run();
            Console.WriteLine("Workbook exported to HTML successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}