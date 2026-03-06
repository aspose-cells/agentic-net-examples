using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";

        string aiRootUrl = "https://api.aspose.cloud";
        string aiKey = "YOUR_API_KEY";

        var cellsAI = new CellsAI(aiRootUrl, aiKey);
        string summary = cellsAI.SpreadsheetSummarize(inputPath);

        var summaryWorkbook = new Workbook();
        var sheet = summaryWorkbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(summary);

        var style = sheet.Cells["A1"].GetStyle();
        style.IsTextWrapped = true;
        sheet.Cells["A1"].SetStyle(style);

        string outputPath = "summary.xlsx";
        summaryWorkbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine($"Summary saved to '{outputPath}'.");
    }
}