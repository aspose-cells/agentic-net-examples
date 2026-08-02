// Title: C# – Export a Single‑Sheet Workbook to HTML with External CSS Folder using Aspose.Cells
// Description: Creates a one‑worksheet workbook, adds sample cells, configures HtmlSaveOptions (ExportWorksheetCSSSeparately = true, CreateDirectory = true), saves the file to a desktop folder, and generates a separate CSS directory (Workbook_files) alongside the HTML page.
// Keywords: Aspose.Cells HTML export | ExportWorksheetCSSSeparately | C# HtmlSaveOptions | external CSS folder Aspose | CreateDirectory option | save Excel as HTML .NET | Workbook_files folder | desktop export path | Aspose.Cells sample code | HTML with separate stylesheet
// Common Searches: Aspose.Cells export HTML separate CSS folder | C# HtmlSaveOptions ExportWorksheetCSSSeparately example | How to create CSS directory when saving Excel as HTML | Save Excel workbook to HTML with external stylesheet .NET | Aspose.Cells HTML export to desktop folder
// Developer Intent: Produce an HTML representation of a workbook while placing its stylesheet in an independent folder.
// Use Cases: Web publishing of Excel reports where CSS needs to be cached separately | Generating static documentation from spreadsheets for intranet sites | Building a desktop tool that converts user‑selected Excel files to HTML with ready‑to‑deploy assets | Automating batch conversion of financial models to HTML for client portals
// AI Prompts: Generate C# code that uses Aspose.Cells to save a workbook as HTML with ExportWorksheetCSSSeparately enabled and ensures the CSS folder is created automatically. | Describe the role of HtmlSaveOptions.CreateDirectory and ExportWorksheetCSSSeparately when exporting Excel to HTML. | Show how to modify the example to export each worksheet of a multi‑sheet workbook to its own HTML file with distinct CSS directories. | Provide a PowerShell script that calls the compiled program to batch‑process multiple Excel files into HTML with separate CSS folders.

using System;
using System.IO;
using Aspose.Cells;

// Creates a one‑worksheet workbook, adds sample cells, configures HtmlSaveOptions (ExportWorksheetCSSSeparately = true, CreateDirectory = true), saves the file to a desktop folder, and generates a separate CSS directory (Workbook_files) alongside the HTML page.
class ExportHtmlWithSeparateCss
{
    static void Main()
    {
        // Create a new workbook with a single worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Sheet1";

        // Add some sample data
        sheet.Cells["A1"].PutValue("Hello World");
        sheet.Cells["B2"].PutValue(12345);

        // Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // Export the worksheet CSS to a separate file
            ExportWorksheetCSSSeparately = true,
            // Automatically create the output directory if it does not exist
            CreateDirectory = true
        };

        // Define the output folder (e.g., Desktop\HtmlExport) and ensure it exists
        string outputFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "HtmlExport");
        Directory.CreateDirectory(outputFolder);

        // Define the full path for the HTML file
        string htmlFilePath = Path.Combine(outputFolder, "Workbook.html");

        // Save the workbook as HTML using the configured options
        workbook.Save(htmlFilePath, saveOptions);

        // The CSS file will be placed in a folder named "Workbook_files" next to the HTML file
        string cssFolder = Path.Combine(outputFolder, "Workbook_files");
        Console.WriteLine($"HTML saved to: {htmlFilePath}");
        Console.WriteLine($"Separate CSS files are in: {cssFolder}");
    }
}
