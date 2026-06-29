using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the merged workbook from disk
        string inputPath = "mergedWorkbook.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.SaveAsSingleFile = true;          // generate a single HTML file
        htmlOptions.ShowAllSheets = true;             // include all worksheets
        htmlOptions.ExportActiveWorksheetOnly = false; // export the whole workbook
        htmlOptions.ExportWorkbookProperties = true;  // keep workbook properties
        htmlOptions.HtmlVersion = HtmlVersion.Html5;  // use HTML5 standard

        // Define the output HTML file path
        string outputPath = "mergedWorkbook.html";

        // Save the workbook as an HTML file
        workbook.Save(outputPath, htmlOptions);

        Console.WriteLine($"HTML representation saved to: {Path.GetFullPath(outputPath)}");
    }
}