using System;
using Aspose.Cells;

namespace AsposeCellsFilePathProviderDemo
{
    // Custom implementation of IFilePathProvider.
    // Returns a full file name for each worksheet when exporting to HTML separately.
    public class CustomFilePathProvider : IFilePathProvider
    {
        public string GetFullName(string sheetName)
        {
            // Example: store each worksheet HTML in a "Sheets" subfolder.
            // Adjust the path format as needed for your environment.
            return $"Sheets\\{sheetName}.html";
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Path to the source XLSX file.
            string sourceXlsx = "input.xlsx";

            // Load the workbook from the XLSX file.
            Workbook workbook = new Workbook(sourceXlsx);

            // Configure HTML save options and assign the custom file path provider.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.FilePathProvider = new CustomFilePathProvider();

            // Save the workbook as HTML. The main HTML file will reference separate sheet files.
            string outputHtml = "output.html";
            workbook.Save(outputHtml, saveOptions);

            Console.WriteLine($"Workbook '{sourceXlsx}' saved to HTML with custom file paths.");
        }
    }
}