using System;
using Aspose.Cells;

namespace AsposeCellsExportPathDemo
{
    // Custom implementation of IFilePathProvider.
    // This class determines the file name (and optionally folder) for each worksheet
    // when the workbook is saved as separate HTML files.
    public class CustomFilePathProvider : IFilePathProvider
    {
        // The folder where the worksheet HTML files will be placed.
        private readonly string _folder;

        public CustomFilePathProvider(string folder)
        {
            _folder = folder;
        }

        // Returns the full path for a given worksheet name.
        // Example: if folder = "Sheets" and sheetName = "Sheet1",
        // the method returns "Sheets/Sheet1.html".
        public string GetFullName(string sheetName)
        {
            // Ensure the folder path ends with a directory separator.
            string normalizedFolder = _folder.EndsWith("/") || _folder.EndsWith("\\")
                ? _folder
                : _folder + System.IO.Path.DirectorySeparatorChar;

            // Combine folder and file name.
            return $"{normalizedFolder}{sheetName}.html";
        }
    }

    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook.
            // Replace "input.xlsx" with the actual path to your source file.
            Workbook workbook = new Workbook("input.xlsx");

            // Create HtmlSaveOptions to control HTML export behavior.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Set the custom file path provider.
            // All worksheet HTML files will be saved under the "ExportedSheets" folder.
            htmlOptions.FilePathProvider = new CustomFilePathProvider("ExportedSheets");

            // Optional: use full path links in the main HTML file.
            // When true, the generated links will contain the absolute path.
            // Set to false for relative links.
            htmlOptions.IsFullPathLink = false;

            // Save the workbook as HTML.
            // The main HTML file will be "output.html" and individual worksheet files
            // will be placed according to the logic in CustomFilePathProvider.
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook exported to HTML with custom worksheet paths.");
        }
    }
}