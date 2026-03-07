using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options with a custom file path provider
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.FilePathProvider = new CustomFilePathProvider();

        // Save the workbook to HTML; each worksheet will be exported using the provider
        workbook.Save("output.html", saveOptions);
    }

    // Custom implementation of IFilePathProvider
    private class CustomFilePathProvider : IFilePathProvider
    {
        // Returns the file name for a worksheet when exporting to HTML separately
        public string GetFullName(string sheetName)
        {
            // Example: "Sheet1.html", "DataSheet.html", etc.
            return $"{sheetName}.html";
        }
    }
}