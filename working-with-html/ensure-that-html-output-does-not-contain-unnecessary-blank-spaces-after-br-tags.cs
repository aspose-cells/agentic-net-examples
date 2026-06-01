using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a cell value that contains a line break.
        // When saved to HTML Aspose.Cells converts the line break to a <br> tag.
        worksheet.Cells["A1"].PutValue("First line\n   Second line with spaces");

        // Save the workbook to HTML using default HtmlSaveOptions.
        string htmlFilePath = "output.html";
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        workbook.Save(htmlFilePath, htmlOptions);

        // Load the generated HTML, remove any spaces (or other whitespace) that appear
        // directly after a <br> tag, and overwrite the file.
        string htmlContent = File.ReadAllText(htmlFilePath);
        string cleanedHtml = Regex.Replace(htmlContent, @"<br>\s+", "<br>");
        File.WriteAllText(htmlFilePath, cleanedHtml);

        Console.WriteLine("HTML file saved and redundant spaces after <br> tags removed.");
    }
}