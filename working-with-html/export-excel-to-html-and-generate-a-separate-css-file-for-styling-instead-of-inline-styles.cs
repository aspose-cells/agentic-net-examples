using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

class ExportExcelToHtmlWithSeparateCss
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John Doe");
        sheet.Cells["B2"].PutValue(30);

        // Apply a style to demonstrate CSS generation
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.Font.Color = Color.Blue;
        sheet.Cells["A1"].SetStyle(headerStyle);
        sheet.Cells["B1"].SetStyle(headerStyle);

        // Configure HTML save options to export CSS in a separate file
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportWorksheetCSSSeparately = true;

        // Define output directory (e.g., Desktop\HtmlExport) and ensure it exists
        string outputDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "HtmlExport");
        Directory.CreateDirectory(outputDir);

        // Save the workbook as HTML; a separate .css file will be created in the same folder
        string htmlPath = Path.Combine(outputDir, "Workbook.html");
        workbook.Save(htmlPath, saveOptions);

        Console.WriteLine("HTML file saved to: " + htmlPath);
        Console.WriteLine("Separate CSS file generated alongside the HTML.");
    }
}