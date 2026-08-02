// Title: Aspose.Cells for .NET – Set HtmlSaveOptions.PageTitle to Workbook File Name (C#)
// Description: Loads an Excel workbook with Aspose.Cells, sets HtmlSaveOptions.PageTitle to the source file name (without extension), and saves the workbook as HTML so the generated <title> tag mirrors the original Excel name.
// Keywords: Aspose.Cells | HtmlSaveOptions | PageTitle | C# HTML export | Excel to HTML conversion | dynamic HTML title | set page title from filename | Aspose.Cells .NET example | save workbook as HTML | HTML title SEO
// Common Searches: Aspose.Cells set HtmlSaveOptions PageTitle | C# set HTML title from Excel filename | How to export Excel to HTML with custom title using Aspose.Cells | HtmlSaveOptions PageTitle property example | Convert .xlsx to .html with filename as title | Aspose.Cells .NET HTML export SEO title
// Developer Intent: Assign the workbook’s file name as the HTML <title> when saving to HTML.
// Use Cases: Create SEO‑friendly HTML reports where each page’s title reflects the originating Excel file. | Build a batch converter that processes a folder of .xlsx files, outputting .html files with matching titles automatically. | Integrate into a web service that streams Excel content as HTML, displaying the original workbook name in the browser tab for better user navigation.
// AI Prompts: Show a C# snippet that loads an Excel file with Aspose.Cells, sets HtmlSaveOptions.PageTitle to the file name without extension, and saves it as .html. | Explain how to batch‑convert multiple Excel workbooks to HTML using Aspose.Cells while dynamically assigning each page’s title from its source filename. | Provide step‑by‑step guidance for configuring HtmlSaveOptions.PageTitle for SEO‑optimized HTML output in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook with Aspose.Cells, sets HtmlSaveOptions.PageTitle to the source file name (without extension), and saves the workbook as HTML so the generated <title> tag mirrors the original Excel name.
class Program
{
    static void Main()
    {
        // Path to the source Excel workbook
        string excelPath = "sample.xlsx";

        // Load the workbook from the file system
        Workbook workbook = new Workbook(excelPath);

        // Create HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Set the HTML page title to the workbook's file name (without extension)
        saveOptions.PageTitle = Path.GetFileNameWithoutExtension(excelPath);

        // Determine the output HTML file path (same name, .html extension)
        string htmlPath = Path.ChangeExtension(excelPath, ".html");

        // Save the workbook as HTML using the configured options
        workbook.Save(htmlPath, saveOptions);
    }
}
