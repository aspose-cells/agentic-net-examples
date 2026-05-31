using System;
using System.Globalization;
using Aspose.Cells;

class HtmlToExcelFrench
{
    static void Main()
    {
        // Input HTML file and output Excel file paths
        string htmlPath = "input.html";
        string excelPath = "output.xlsx";

        // Configure load options for HTML format and set French culture (decimal separator = ',')
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
        loadOptions.CultureInfo = new CultureInfo("fr-FR");

        // Load the HTML file into a workbook using the specified load options
        Workbook workbook = new Workbook(htmlPath, loadOptions);

        // Apply French regional settings to the workbook (affects formatting, formulas, etc.)
        workbook.Settings.Region = CountryCode.France;
        workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

        // Save the workbook as an Excel file (XLSX)
        workbook.Save(excelPath, SaveFormat.Xlsx);
    }
}