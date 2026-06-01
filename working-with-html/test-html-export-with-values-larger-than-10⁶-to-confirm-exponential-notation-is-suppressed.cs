using System;
using System.IO;
using Aspose.Cells;

class HtmlExportLargeNumberDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert a large numeric value (> 10^6)
        worksheet.Cells["A1"].PutValue(1234567890);

        // Apply a number format that forces plain notation (no scientific format)
        Style style = worksheet.Cells["A1"].GetStyle();
        style.Custom = "0";               // Integer format without exponent
        worksheet.Cells["A1"].SetStyle(style);

        // Configure HTML save options (default settings are sufficient)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportFormula = false; // Not needed for this test

        // Save the workbook as HTML
        string htmlFile = "LargeNumber.html";
        workbook.Save(htmlFile, htmlOptions);

        // Load the generated HTML and verify that exponential notation is not present
        string htmlContent = File.ReadAllText(htmlFile);
        bool hasExponential = htmlContent.Contains("E+");
        bool hasPlainNumber = htmlContent.Contains("1234567890");

        Console.WriteLine("Exponential notation present: " + hasExponential);
        Console.WriteLine("Plain number present: " + hasPlainNumber);
    }
}