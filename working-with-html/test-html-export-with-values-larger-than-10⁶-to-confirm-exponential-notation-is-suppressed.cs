// Title: Suppress scientific notation when exporting numbers >10⁶ to HTML with Aspose.Cells for .NET
// Description: Demonstrates how to place a value larger than one million in a worksheet, apply a custom "0" number format, save the workbook as HTML using HtmlSaveOptions, and verify that the generated HTML contains the plain integer instead of exponential (E+) notation.
// Keywords: Aspose.Cells HTML export large numbers | prevent scientific notation Aspose.Cells | custom number format 0 Aspose.Cells | C# HtmlSaveOptions suppress exponential notation | verify HTML output Aspose.Cells
// Common Searches: Aspose.Cells export large integer to HTML without scientific notation | How to stop HTML output from showing 1E+09 in Aspose.Cells .NET | C# save workbook as HTML plain number format | Suppress exponential notation in Aspose.Cells HTML export
// Developer Intent: Export cells containing values over 1,000,000 as plain integers in HTML, ensuring no scientific notation appears.
// Use Cases: Financial dashboards where full monetary figures must be displayed in HTML reports. | Invoice generators that list large item IDs or quantities without scientific notation. | Automated regression tests that confirm HTML export renders large numeric cells correctly.
// AI Prompts: Generate C# code with Aspose.Cells that writes a number >10⁶ to a cell, applies a "0" format, saves to HTML, and checks the file for absence of "E+" strings. | Create a unit test in C# that asserts Aspose.Cells HTML export does not use exponential notation for large numbers. | Explain the role of HtmlSaveOptions and cell style settings in preventing scientific notation during HTML conversion with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to place a value larger than one million in a worksheet, apply a custom "0" number format, save the workbook as HTML using HtmlSaveOptions, and verify that the generated HTML contains the plain integer instead of exponential (E+) notation.
class HtmlExportLargeNumberDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert a large numeric value (greater than 10⁶)
        worksheet.Cells["A1"].PutValue(1234567890);

        // Apply a number format that forces plain integer representation
        Style intStyle = workbook.CreateStyle();
        intStyle.Custom = "0"; // No decimal places, no scientific notation
        worksheet.Cells["A1"].SetStyle(intStyle);

        // Configure HTML save options (default settings are sufficient)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.CalculateFormula = true; // Ensure any formulas are evaluated

        // Save the workbook as HTML
        string htmlFile = "LargeNumber.html";
        workbook.Save(htmlFile, htmlOptions);

        // Load the generated HTML and verify that exponential notation is not present
        string htmlContent = File.ReadAllText(htmlFile);
        bool hasExponentialNotation = htmlContent.Contains("E+") || htmlContent.Contains("e+");
        Console.WriteLine("Exponential notation present: " + hasExponentialNotation);
        Console.WriteLine("HTML file saved at: " + Path.GetFullPath(htmlFile));
    }
}
