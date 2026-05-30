using System;
using System.IO;
using Aspose.Cells;

class ExportHtmlSeparateCss
{
    static void Main()
    {
        // Create a new workbook with a single worksheet and add some data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Sheet1";
        sheet.Cells["A1"].PutValue("Sample Text");
        sheet.Cells["B2"].PutValue(12345);

        // Configure HTML save options to export worksheet CSS separately
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportWorksheetCSSSeparately = true, // Enable separate CSS files per worksheet
            CreateDirectory = true               // Auto‑create output directories if they don't exist
        };

        // Define the output folder (e.g., Desktop\HtmlExport) and HTML file name
        string outputFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "HtmlExport");

        string htmlFilePath = Path.Combine(outputFolder, "Workbook.html");

        // Save the workbook as HTML; a "css" subfolder will be created automatically
        workbook.Save(htmlFilePath, saveOptions);

        Console.WriteLine($"HTML file saved to: {htmlFilePath}");
        Console.WriteLine($"CSS folder created at: {Path.Combine(outputFolder, "css")}");
    }
}