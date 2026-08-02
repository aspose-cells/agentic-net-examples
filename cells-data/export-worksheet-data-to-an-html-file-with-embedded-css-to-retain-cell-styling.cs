using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

class ExportWorksheetToHtml
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Age");
        worksheet.Cells["A2"].PutValue("John Doe");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["A3"].PutValue("Jane Smith");
        worksheet.Cells["B3"].PutValue(28);

        // Apply a simple style to the header row to demonstrate CSS retention
        Style headerStyle = worksheet.Cells["A1"].GetStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.Font.Color = Color.White;
        headerStyle.ForegroundColor = Color.DarkBlue;
        headerStyle.Pattern = BackgroundType.Solid;
        worksheet.Cells["A1"].SetStyle(headerStyle);
        worksheet.Cells["B1"].SetStyle(headerStyle);

        // Configure HTML save options to embed CSS within the HTML file
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportWorksheetCSSSeparately = false; // embed CSS
        saveOptions.DisableCss = false;                  // use CSS (not inline only)

        // Define output path (e.g., Desktop)
        string outputPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "WorksheetExport.html");

        // Save the workbook as an HTML file with embedded CSS
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine("Worksheet exported to HTML with embedded CSS at:");
        Console.WriteLine(outputPath);
    }
}