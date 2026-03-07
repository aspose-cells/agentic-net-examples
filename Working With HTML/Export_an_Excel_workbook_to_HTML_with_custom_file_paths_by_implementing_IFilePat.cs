using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Custom implementation of IFilePathProvider to control the file names of
    // individual worksheet HTML files when the workbook is saved as separate HTML pages.
    public class CustomFilePathProvider : IFilePathProvider
    {
        // Returns a custom file name based on the worksheet name.
        // For example, a sheet named "Sheet1" will be saved as "custom_Sheet1.html".
        public string GetFullName(string sheetName)
        {
            return $"custom_{sheetName}.html";
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Load the source Excel workbook (XLSX format).
            // Replace "input.xlsx" with the actual path to your source file.
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options and assign the custom file path provider.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.FilePathProvider = new CustomFilePathProvider();

            // Save the workbook as HTML. Each worksheet will be saved using the
            // file name returned by CustomFilePathProvider.
            // Replace "output.html" with the desired main HTML file name.
            workbook.Save("output.html", saveOptions);
        }
    }
}