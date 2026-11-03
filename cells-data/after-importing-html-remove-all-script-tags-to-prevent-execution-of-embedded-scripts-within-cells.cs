// Title: Remove <script> tags from HTML content imported into an Aspose.Cells workbook using C#
// Common Searches: C# Aspose.Cells remove script tags from HTML after loading workbook | how to strip <script> elements from cells when importing HTML with Aspose.Cells | sanitize HTML content in Aspose.Cells workbook to prevent JavaScript execution | regex approach to delete script tags from cell values in .NET Aspose.Cells | post‑import HTML cleaning for Excel cells using Aspose.Cells and C# | prevent embedded scripts in Excel cells after HtmlLoadOptions import
// Tags: script tag removal Aspose.Cells | HTML content sanitization workbook cells C# | post‑load cell cleaning Aspose.Cells | regex based HTML cleanup Aspose.Cells | prevent JavaScript execution Excel .NET | HtmlLoadOptions cell value sanitization

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an HTML file into an Aspose.Cells Workbook with HtmlLoadOptions, scans every worksheet and cell, and uses a case‑insensitive regular expression to strip any <script>...</script> fragments from string cell values before saving the workbook as an XLSX file, ensuring no embedded scripts can execute.
class RemoveScriptTagsFromHtml
{
    static void Main()
    {
        // Path to the HTML file to import
        string htmlPath = "input.html";

        // Verify that the HTML file exists
        if (!File.Exists(htmlPath))
        {
            Console.WriteLine($"Error: File \"{htmlPath}\" not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load the HTML content into a workbook using HtmlLoadOptions
            workbook = new Workbook(htmlPath, new HtmlLoadOptions());
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading HTML file: {ex.Message}");
            return;
        }

        // Regular expression to match <script> tags (case‑insensitive, singleline)
        Regex scriptRegex = new Regex(@"<script\b[^>]*>.*?</script>", RegexOptions.IgnoreCase | RegexOptions.Singleline);

        // Iterate through all worksheets and cells to clean script tags
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            var cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.Type == CellValueType.IsString)
                    {
                        string original = cell.StringValue;
                        // Remove any <script>...</script> fragments
                        string cleaned = scriptRegex.Replace(original, string.Empty);
                        if (!original.Equals(cleaned))
                        {
                            cell.PutValue(cleaned);
                        }
                    }
                }
            }
        }

        // Save the cleaned workbook to a new file
        try
        {
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved successfully as output.xlsx.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
