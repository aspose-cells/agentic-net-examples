using System;
using Aspose.Cells;

namespace AsposeCellsFilePathProviderDemo
{
    // Custom implementation of IFilePathProvider.
    // Returns a file name for each worksheet when exporting to HTML separately.
    public class CustomFilePathProvider : IFilePathProvider
    {
        // The method required by the interface.
        // It receives the worksheet name and should return the full path (or relative file name) for that sheet.
        public string GetFullName(string sheetName)
        {
            // Example: place each sheet in its own HTML file named "<sheetName>.html"
            // Adjust the path as needed (e.g., include a folder).
            return $"{sheetName}.html";
        }
    }

    public static class Program
    {
        // Demonstrates loading an XLSX workbook and saving it to HTML using the custom file path provider.
        public static void Run()
        {
            // Path to the source Excel file.
            string sourcePath = "input.xlsx";

            // Load the workbook from the XLSX file.
            Workbook workbook = new Workbook(sourcePath);

            // Create HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Assign the custom file path provider so that each worksheet is saved to a separate HTML file.
            htmlOptions.FilePathProvider = new CustomFilePathProvider();

            // Save the workbook to HTML.
            // The main HTML file will be created, and additional worksheet files will be generated
            // according to the logic in CustomFilePathProvider.
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook loaded from XLSX and saved to HTML with custom file paths.");
        }

        // Entry point required by the .NET runtime.
        public static void Main(string[] args)
        {
            Run();
        }
    }
}