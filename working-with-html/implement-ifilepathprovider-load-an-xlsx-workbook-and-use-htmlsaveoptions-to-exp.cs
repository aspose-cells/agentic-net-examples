using System;
using Aspose.Cells;

class CustomFilePathProvider : IFilePathProvider
{
    // Generates a custom file name for each worksheet when exporting to HTML.
    public string GetFullName(string sheetName)
    {
        // Example: "custom_Sheet1.html"
        return $"custom_{sheetName}.html";
    }
}

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook (adjust the path as needed).
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Set up HTML save options and assign the custom file path provider.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.FilePathProvider = new CustomFilePathProvider();

        // Export the workbook to HTML. Separate worksheet files will use the custom names.
        workbook.Save("output.html", saveOptions);
    }
}