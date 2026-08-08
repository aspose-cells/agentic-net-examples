// Title: C# – Remove <script> Tags from HTML Imported into Aspose.Cells Workbook
// Description: Loads an HTML file with Aspose.Cells, walks through every used cell, deletes <script>…</script> blocks from the cell's HtmlString using a case‑insensitive regular expression, and saves the cleaned workbook as XLSX.
// Keywords: Aspose.Cells | .NET | C# | remove script tags | HTML import | sanitize workbook | regex HtmlString | prevent script execution | Excel security | strip JavaScript
// Common Searches: Aspose.Cells remove script tags C# | strip JavaScript from imported HTML cells | sanitize HTML in Excel workbook using Aspose.Cells | prevent script execution in Excel .NET | clean HtmlString Aspose.Cells example
// Developer Intent: Eliminate every <script> element from cells after loading HTML to stop embedded scripts from running.
// Use Cases: Sanitize a workbook generated from untrusted HTML before sharing it with clients. | Integrate the cleaning step into an automated report pipeline that converts web pages to Excel. | Extend the regex to also remove <style> or other unwanted tags while preserving cell formatting.
// AI Prompts: Write C# code that loads an HTML file with Aspose.Cells, removes all <script> tags from each cell's HtmlString using a regular expression, and saves the result as an XLSX file. | Suggest a non‑regex method to strip script elements from HTML imported with Aspose.Cells, such as using an HTML parser or custom load options. | Generate a unit test that verifies script tags are removed from cells after processing with the provided code.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

// Loads an HTML file with Aspose.Cells, walks through every used cell, deletes <script>…</script> blocks from the cell's HtmlString using a case‑insensitive regular expression, and saves the cleaned workbook as XLSX.
class RemoveScriptTagsFromHtmlImport
{
    static void Main()
    {
        // Load the HTML file with default load options
        HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
        Workbook workbook = new Workbook("input.html", loadOptions);

        // Regular expression to strip <script>...</script> tags (case‑insensitive, multiline)
        Regex scriptRegex = new Regex(@"<script\b[^>]*>.*?</script>", RegexOptions.IgnoreCase | RegexOptions.Singleline);

        // Iterate through all worksheets and their used cells
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    string html = cell.HtmlString;

                    if (!string.IsNullOrEmpty(html))
                    {
                        string cleaned = scriptRegex.Replace(html, string.Empty);
                        if (cleaned != html)
                        {
                            cell.HtmlString = cleaned;
                        }
                    }
                }
            }
        }

        // Save the cleaned workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
