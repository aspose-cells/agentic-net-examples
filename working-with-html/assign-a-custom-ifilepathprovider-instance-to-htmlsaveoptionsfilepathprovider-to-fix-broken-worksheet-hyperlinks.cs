using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Custom implementation of IFilePathProvider.
    // Returns a file name for each worksheet when exporting to HTML separately.
    internal class CustomFilePathProvider : IFilePathProvider
    {
        public string GetFullName(string sheetName)
        {
            // Example: place each worksheet in its own HTML file named "<sheetName>.html"
            return $"{sheetName}.html";
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and add two worksheets.
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";
            sheet1.Cells["A1"].PutValue("Hello from Sheet 1");

            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
            sheet2.Cells["A1"].PutValue("Hello from Sheet 2");

            // Add a hyperlink in Sheet1 that points to Sheet2.
            // The link will be resolved correctly when the custom FilePathProvider is used.
            sheet1.Hyperlinks.Add("B1", 1, 1, $"#{sheet2.Name}!A1");

            // Configure HTML save options and assign the custom file path provider.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.FilePathProvider = new CustomFilePathProvider();

            // Save the workbook to HTML. Each worksheet will be saved to its own HTML file
            // using the paths supplied by CustomFilePathProvider.
            workbook.Save("Workbook.html", saveOptions);

            Console.WriteLine("Workbook saved to HTML with custom file paths.");
        }
    }
}