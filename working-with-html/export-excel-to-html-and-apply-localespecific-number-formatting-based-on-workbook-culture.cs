using System;
using System.Globalization;
using Aspose.Cells;

class ExportExcelToHtmlWithLocale
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and its cells
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate cells with numeric values
        cells["A1"].PutValue(1234.56);
        cells["A2"].PutValue(7890.12);

        // Apply a built‑in number format (currency) to the cells
        Style currencyStyle = workbook.CreateStyle();
        currencyStyle.Number = 164; // Currency format
        cells["A1"].SetStyle(currencyStyle);
        cells["A2"].SetStyle(currencyStyle);

        // Set the workbook culture to French (France) so numbers are formatted accordingly
        workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        // Export all data (including styles and formatting)
        htmlOptions.ExportDataOptions = HtmlExportDataOptions.All;

        // Save the workbook as an HTML file using the specified options
        workbook.Save("output.html", htmlOptions);
    }
}