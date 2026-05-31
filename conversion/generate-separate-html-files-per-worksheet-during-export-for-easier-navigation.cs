using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Rename the default sheet and add two more worksheets
        workbook.Worksheets[0].Name = "Summary";
        Worksheet dataSheet = workbook.Worksheets.Add("Data");
        Worksheet reportSheet = workbook.Worksheets.Add("Report");

        // Populate each worksheet with sample data
        workbook.Worksheets["Summary"].Cells["A1"].PutValue("This is the Summary sheet.");
        dataSheet.Cells["A1"].PutValue("This is the Data sheet.");
        reportSheet.Cells["A1"].PutValue("This is the Report sheet.");

        // Configure HTML save options to generate a separate HTML file for each worksheet
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.SaveAsSingleFile = false;               // Do not combine sheets into one file
        saveOptions.ExportActiveWorksheetOnly = false;      // Export all worksheets
        saveOptions.FilePathProvider = new CustomFilePathProvider(); // Custom file naming

        // Save the workbook; Aspose.Cells will create individual HTML files per sheet
        workbook.Save("Workbook.html", saveOptions);
    }

    // Implements IFilePathProvider to control the file name of each exported worksheet
    class CustomFilePathProvider : IFilePathProvider
    {
        public string GetFullName(string sheetName)
        {
            // Each worksheet will be saved as <SheetName>.html
            return $"{sheetName}.html";
        }
    }
}