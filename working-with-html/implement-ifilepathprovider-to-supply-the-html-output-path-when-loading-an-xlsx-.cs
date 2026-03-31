using System;
using Aspose.Cells;

namespace AsposeCellsFilePathProviderDemo
{
    // Custom implementation of IFilePathProvider.
    // Returns a full file name for each worksheet when exporting to HTML separately.
    public class CustomFilePathProvider : IFilePathProvider
    {
        // Example: place each sheet HTML in a subfolder named "Sheets".
        public string GetFullName(string sheetName)
        {
            // Ensure the folder exists when saving (Aspose will handle the path creation).
            // Return a relative path that will be used in the main HTML file.
            return $"Sheets/{sheetName}.html";
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source XLSX workbook.
            string inputPath = "input.xlsx";

            // Load the workbook from the XLSX file.
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            // Assign the custom file path provider.
            saveOptions.FilePathProvider = new CustomFilePathProvider();

            // Save the workbook as HTML. The main HTML file will be "output.html"
            // and each worksheet will be saved to the paths returned by the provider.
            string outputHtml = "output.html";
            workbook.Save(outputHtml, saveOptions);

            Console.WriteLine($"Workbook saved to HTML with custom file paths. Main file: {outputHtml}");
        }
    }
}