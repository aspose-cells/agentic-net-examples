// Title: Suppress Scientific Notation in Aspose.Cells HTML Export and Verify Numeric Accuracy (C#)
// Description: Creates a workbook with a tiny number, saves it to HTML using default settings, applies a custom decimal format to hide scientific notation, saves a second HTML file, extracts the displayed values with a regex, and confirms that the formatted output matches the original number.
// Keywords: Aspose.Cells HTML export | suppress scientific notation | custom number format C# | numeric accuracy verification | regex cell value extraction | Excel to HTML conversion | C# Aspose.Cells example | global development
// Common Searches: Aspose.Cells prevent scientific notation in HTML export | compare default and custom number formats in Aspose.Cells HTML | extract cell text from Aspose.Cells generated HTML | C# verify numeric display after HTML conversion | how to format small numbers in Aspose.Cells HTML output
// Developer Intent: The developer wants to ensure that applying a custom decimal format removes exponential notation from the HTML export and that the displayed value exactly matches the original numeric value.
// Use Cases: Generate HTML from an Excel workbook while keeping very small numbers in plain decimal form. | Automated test that compares default HTML output with a formatted version to detect numeric representation changes. | Validate that the cell value extracted from the suppressed‑notation HTML equals the expected formatted string.
// AI Prompts: Write C# code to compare two Aspose.Cells HTML files and highlight numeric differences. | Suggest a reliable alternative to regex for extracting cell values from Aspose.Cells HTML output. | Explain how to apply a workbook‑wide custom number format to eliminate scientific notation before saving to HTML.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsHtmlComparison
{
    // Creates a workbook with a tiny number, saves it to HTML using default settings, applies a custom decimal format to hide scientific notation, saves a second HTML file, extracts the displayed values with a regex, and confirms that the formatted output matches the original number.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put a very small number that Excel would normally display in exponential notation
            double smallNumber = 0.000000123456789;
            cells["A1"].PutValue(smallNumber);

            // -----------------------------------------------------------------
            // Export HTML with default settings (exponential notation may appear)
            // -----------------------------------------------------------------
            HtmlSaveOptions defaultOptions = new HtmlSaveOptions();
            string defaultHtmlPath = "default.html";
            workbook.Save(defaultHtmlPath, defaultOptions);

            // -----------------------------------------------------------------
            // Apply a custom number format to suppress exponential notation
            // Example format: show up to 15 decimal places without scientific format
            // -----------------------------------------------------------------
            Style customStyle = cells["A1"].GetStyle();
            customStyle.Custom = "0.####################"; // Plain decimal format
            cells["A1"].SetStyle(customStyle);

            // Export HTML after applying the custom format
            HtmlSaveOptions suppressedOptions = new HtmlSaveOptions();
            string suppressedHtmlPath = "suppressed.html";
            workbook.Save(suppressedHtmlPath, suppressedOptions);

            // -----------------------------------------------------------------
            // Load both HTML files as strings
            // -----------------------------------------------------------------
            string htmlDefault = File.ReadAllText(defaultHtmlPath);
            string htmlSuppressed = File.ReadAllText(suppressedHtmlPath);

            // -----------------------------------------------------------------
            // Extract the displayed cell value from each HTML using a simple regex.
            // The cell value is typically inside a <td> element.
            // -----------------------------------------------------------------
            string pattern = @"<td[^>]*>(.*?)</td>";
            string valueDefault = ExtractFirstMatch(htmlDefault, pattern);
            string valueSuppressed = ExtractFirstMatch(htmlSuppressed, pattern);

            // -----------------------------------------------------------------
            // Output the comparison results
            // -----------------------------------------------------------------
            Console.WriteLine("Original numeric value: " + smallNumber);
            Console.WriteLine("HTML with default settings   : " + valueDefault);
            Console.WriteLine("HTML with exponential suppression: " + valueSuppressed);

            // Simple accuracy check: compare the suppressed value with the original number formatted as plain text
            string expectedPlain = smallNumber.ToString("0.####################");
            bool isAccurate = string.Equals(valueSuppressed, expectedPlain, StringComparison.Ordinal);
            Console.WriteLine("Suppressed HTML matches expected plain format: " + isAccurate);
        }

        // Helper method to get the first captured group from a regex match
        private static string ExtractFirstMatch(string input, string pattern)
        {
            Match match = Regex.Match(input, pattern, RegexOptions.Singleline);
            if (match.Success && match.Groups.Count > 1)
            {
                // Remove any HTML tags that might be inside the cell (e.g., <span>)
                string inner = Regex.Replace(match.Groups[1].Value, "<.*?>", string.Empty);
                return inner.Trim();
            }
            return string.Empty;
        }
    }
}
