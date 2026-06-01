using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class HtmlToExcelWithPassword
{
    static void Main()
    {
        // Paths for the source HTML file and the target Excel file
        string htmlFilePath = "input.html";
        string excelFilePath = "output.xlsx";

        // Password to protect the resulting workbook
        string workbookPassword = "mySecretPassword";

        // Load the HTML file into a Workbook using LoadOptions for HTML format
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
        Workbook workbook = new Workbook(htmlFilePath, loadOptions);

        // Apply password protection to the workbook
        workbook.Settings.Password = workbookPassword;

        // Save the workbook as an Excel file (XLSX format)
        workbook.Save(excelFilePath, SaveFormat.Xlsx);

        Console.WriteLine("HTML has been converted to Excel and saved with password protection.");
    }
}