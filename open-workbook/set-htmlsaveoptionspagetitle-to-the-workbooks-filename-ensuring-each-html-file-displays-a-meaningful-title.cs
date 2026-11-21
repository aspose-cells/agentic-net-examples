// Title: Assign workbook filename as HTML page title when converting Excel to HTML with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set HTML page title to Excel file name during save | How to use HtmlSaveOptions.PageTitle to display workbook name in generated HTML | Save Excel workbook as HTML with dynamic title based on source filename using Aspose.Cells | Aspose.Cells HtmlSaveOptions custom page title from Path.GetFileNameWithoutExtension | Generate HTML from Excel with title tag matching original workbook name in .NET
// Tags: HtmlSaveOptions page title | dynamic HTML title for Excel conversion | Aspose.Cells set page title from file name | C# extract workbook filename for HTML title | save workbook as HTML with custom title

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, extracts the file name without its extension, assigns that string to HtmlSaveOptions.PageTitle, and saves the workbook as an HTML file so the resulting page’s <title> element reflects the original workbook name.
class Program
{
    static void Main()
    {
        // Path to the source Excel workbook
        string workbookPath = @"C:\Data\SampleWorkbook.xlsx";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(workbookPath);

        // Extract the file name without extension to use as the HTML page title
        string pageTitle = Path.GetFileNameWithoutExtension(workbookPath);

        // Configure HTML save options and set the page title
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.PageTitle = pageTitle; // Ensures the HTML file displays a meaningful title

        // Save the workbook as an HTML file with the specified options
        string htmlOutputPath = @"C:\Data\SampleWorkbook.html";
        workbook.Save(htmlOutputPath, htmlOptions);
    }
}
