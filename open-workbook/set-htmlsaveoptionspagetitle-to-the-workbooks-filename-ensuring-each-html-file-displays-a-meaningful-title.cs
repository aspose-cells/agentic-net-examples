// Title: How to set HtmlSaveOptions.PageTitle to the workbook filename when converting Excel to HTML with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set HTML page title to Excel file name | HtmlSaveOptions.PageTitle example for converting .xlsx to .html | How to use Path.GetFileNameWithoutExtension for HTML title in Aspose.Cells | Export Excel workbook as HTML with custom browser tab title using Aspose.Cells | C# code to assign workbook filename as HTML page title in Aspose.Cells
// Tags: HtmlSaveOptions.PageTitle property | Excel to HTML conversion Aspose.Cells | C# set HTML title from workbook filename | Path.GetFileNameWithoutExtension usage Aspose.Cells | custom page title for exported HTML

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, creates HtmlSaveOptions, sets the PageTitle to the workbook's file name without its extension using Path.GetFileNameWithoutExtension, and saves the workbook as an HTML file so the generated page displays the original filename as its title.
class HtmlExportWithTitle
{
    static void Main()
    {
        // Path to the source Excel file
        string excelPath = @"C:\Data\Report.xlsx";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(excelPath);

        // Create HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

        // Set the page title to the workbook's file name (without extension)
        htmlOptions.PageTitle = Path.GetFileNameWithoutExtension(excelPath);

        // Define the output HTML file path
        string htmlPath = @"C:\Data\Report.html";

        // Save the workbook as HTML using the configured options
        workbook.Save(htmlPath, htmlOptions);
    }
}
