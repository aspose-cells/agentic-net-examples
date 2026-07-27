// Title: Validate TableCssId Prefix in HTML Exported by Aspose.Cells for .NET (C#)
// Description: A C# sample that creates a workbook, assigns a custom TableCssId via HtmlSaveOptions, saves the file as HTML, reads the output, and uses a regular expression to confirm the <table> element contains the expected TableCssId value.
// Keywords: Aspose.Cells | HtmlSaveOptions | TableCssId | C# HTML export | regex validation | custom CSS ID | unit testing | CI/CD verification | .NET spreadsheet to HTML | table attribute check
// Common Searches: Aspose.Cells verify TableCssId in exported HTML | C# regex check table id in Aspose.Cells HTML output | How to test custom TableCssId prefix with Aspose.Cells | Validate HTML export from Aspose.Cells .NET | Check table CSS identifier after saving workbook as HTML
// Developer Intent: Confirm that the HTML produced by Aspose.Cells includes the specified TableCssId on the generated <table> element.
// Use Cases: Automated unit test that asserts the presence of a custom TableCssId after HTML export. | CI/CD pipeline step to ensure spreadsheet‑to‑HTML conversions use correct CSS identifiers. | Debugging styling problems caused by missing or incorrect table IDs in web pages.
// AI Prompts: Generate a C# method that loads an Aspose.Cells HTML file and returns true if the TableCssId attribute matches a given prefix. | Create an NUnit test that saves a workbook with a custom TableCssId and verifies the attribute exists in the resulting HTML. | Write code using HtmlAgilityPack to locate the <table> element and validate its TableCssId attribute instead of using regex.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// A C# sample that creates a workbook, assigns a custom TableCssId via HtmlSaveOptions, saves the file as HTML, reads the output, and uses a regular expression to confirm the <table> element contains the expected TableCssId value.
class ValidateTableCssId
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Age");
        worksheet.Cells["A2"].PutValue("John");
        worksheet.Cells["B2"].PutValue(30);

        // Configure HTML save options with a specific TableCssId prefix
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        string tableCssIdPrefix = "custom-table-style";
        saveOptions.TableCssId = tableCssIdPrefix;

        // Save the workbook as HTML
        string htmlFilePath = "output.html";
        workbook.Save(htmlFilePath, saveOptions);

        // Load the generated HTML content
        string htmlContent = File.ReadAllText(htmlFilePath);

        // Validate that the <table> element contains the expected TableCssId attribute
        string pattern = $@"<table[^>]*\bTableCssId\s*=\s*[""']{Regex.Escape(tableCssIdPrefix)}[""']";
        bool isValid = Regex.IsMatch(htmlContent, pattern, RegexOptions.IgnoreCase);

        // Output validation result
        Console.WriteLine(isValid
            ? $"Validation succeeded: <table> contains TableCssId=\"{tableCssIdPrefix}\"."
            : $"Validation failed: <table> does not contain the expected TableCssId=\"{tableCssIdPrefix}\".");
    }
}
