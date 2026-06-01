using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a workbook and add a cell with redundant spaces
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("   This    text   has   redundant   spaces   ");

        // Save the workbook as HTML (original version with spaces)
        string originalHtmlPath = "original.html";
        workbook.Save(originalHtmlPath, new HtmlSaveOptions());

        // Load the generated HTML with DeleteRedundantSpaces enabled
        HtmlLoadOptions loadOptions = new HtmlLoadOptions
        {
            DeleteRedundantSpaces = true
        };

        Workbook cleanedWorkbook;
        using (FileStream fs = new FileStream(originalHtmlPath, FileMode.Open, FileAccess.Read))
        {
            cleanedWorkbook = new Workbook(fs, loadOptions);
        }

        // Save the cleaned HTML
        string cleanedHtmlPath = "cleaned.html";
        cleanedWorkbook.Save(cleanedHtmlPath, new HtmlSaveOptions());

        Console.WriteLine($"Original HTML saved to: {originalHtmlPath}");
        Console.WriteLine($"Cleaned HTML saved to: {cleanedHtmlPath}");
    }
}