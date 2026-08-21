// Title: C# – Set HtmlSaveOptions.PageTitle to Excel filename when converting to HTML with Aspose.Cells
// Description: Load an Excel workbook using Aspose.Cells, assign the source file name to HtmlSaveOptions.PageTitle, and save the workbook as an HTML file so the generated <title> tag matches the original spreadsheet name.
// Keywords: Aspose.Cells | HtmlSaveOptions.PageTitle | C# Excel to HTML | set HTML title from filename | .NET spreadsheet conversion | Excel workbook HTML export | SEO friendly HTML title
// Common Searches: Aspose.Cells set HTML page title from Excel file name | C# HtmlSaveOptions PageTitle example | convert .xlsx to .html with original filename as title | how to change <title> tag when saving Excel as HTML using Aspose | Aspose.Cells .NET HTML export filename title
// Developer Intent: Assign the workbook's file name to HtmlSaveOptions.PageTitle before saving the workbook as an HTML document.
// Use Cases: Generate HTML reports where the browser tab shows the source Excel name for easy identification. | Batch‑process a folder of spreadsheets, automatically using each file's name as the HTML page title. | Create SEO‑optimized web pages from spreadsheets by embedding the original filename in the <title> element.
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, sets HtmlSaveOptions.PageTitle to the workbook's filename, and saves it as HTML. | Provide a script that iterates over all .xlsx files in a directory, converts each to .html with Aspose.Cells, and sets the HTML <title> to the corresponding file name. | Explain the impact of HtmlSaveOptions.PageTitle on the generated HTML and why using the workbook filename improves usability and SEO.

using System;
using System.IO;
using Aspose.Cells;

// Load an Excel workbook using Aspose.Cells, assign the source file name to HtmlSaveOptions.PageTitle, and save the workbook as an HTML file so the generated <title> tag matches the original spreadsheet name.
class Program
{
    static void Main(string[] args)
    {
        // Path to the source Excel file
        string excelPath = "input.xlsx";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(excelPath);

        // Create HTML save options using the default constructor
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Set the HTML page title to the workbook's file name (including extension)
        htmlOptions.PageTitle = Path.GetFileName(excelPath);

        // Determine the output HTML file path (same name with .html extension)
        string htmlPath = Path.ChangeExtension(excelPath, ".html");

        // Save the workbook as an HTML file with the specified page title
        workbook.Save(htmlPath, htmlOptions);

        Console.WriteLine($"HTML file saved to '{htmlPath}' with page title '{htmlOptions.PageTitle}'.");
    }
}
