// Title: Export a workbook to HTML with Aspose.Cells while keeping large numbers displayed as plain text (no scientific notation)
// AI Prompts: Generate C# code that writes a very large double to a cell, applies a custom number format to force integer display, saves the workbook as HTML, and then scans the resulting HTML to confirm the number is not shown in exponential form. | Write a C# program using Aspose.Cells that formats large numeric values with a "0" custom format, exports the worksheet to HTML, reads the HTML file, and validates that no scientific‑notation pattern appears.
// Common Searches: Aspose.Cells export to HTML large double without scientific notation | C# prevent exponential format when saving Excel as HTML using Aspose.Cells | how to verify numeric formatting in generated HTML file Aspose.Cells | custom number format 0 for large numbers Aspose.Cells C#
// Tags: HTML export custom number format Aspose.Cells | suppress scientific notation large numbers Aspose.Cells | validate numeric representation in HTML C# | apply integer format to cell Aspose.Cells | read generated HTML file Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The program creates a new workbook, inserts a very large double into cell A1, sets a custom number format "0" to force plain‑text integer display, saves the workbook as HTML, reads the generated file, and uses a regular expression to ensure the number is not rendered in exponential notation.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define a large number
            double largeNumber = 1234567890123456789.0;

            // Put the large number into cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(largeNumber);

            // Apply a custom number format to force plain‑text representation (no scientific notation)
            Style style = cell.GetStyle();
            style.Custom = "0"; // Display as integer without exponent
            cell.SetStyle(style);

            // Save the workbook as HTML
            string htmlPath = "LargeNumber.html";
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            // The ExportNumericFormat property is not required for this version; default behavior preserves formatting
            workbook.Save(htmlPath, saveOptions);

            // Ensure the HTML file was created before reading
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException($"The file '{htmlPath}' was not found after saving.");

            // Read the generated HTML
            string htmlContent = File.ReadAllText(htmlPath);

            // Verify that the large number appears as plain text (no exponential notation)
            bool containsExponential = Regex.IsMatch(htmlContent, @"E\+|\e\+|\d\.\d+e", RegexOptions.IgnoreCase);
            if (containsExponential)
                Console.WriteLine("Test Failed: Large number is displayed in exponential form.");
            else
                Console.WriteLine("Test Passed: Large number is displayed as plain text.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
