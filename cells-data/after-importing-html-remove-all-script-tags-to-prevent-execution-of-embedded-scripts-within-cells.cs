using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

class RemoveScriptTagsFromHtmlImport
{
    static void Main()
    {
        // Load the HTML file into a workbook
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        Workbook workbook = new Workbook("input.html", loadOptions);

        // Iterate through all worksheets and cells to clean <script> tags
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
                        // Remove any <script>...</script> blocks (case‑insensitive, multiline)
                        string cleaned = Regex.Replace(
                            html,
                            @"<script\b[^>]*>.*?</script>",
                            string.Empty,
                            RegexOptions.IgnoreCase | RegexOptions.Singleline);

                        if (cleaned != html)
                        {
                            cell.HtmlString = cleaned;
                        }
                    }
                }
            }
        }

        // Optional: ensure no macros are present
        workbook.RemoveMacro();

        // Save the cleaned workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}