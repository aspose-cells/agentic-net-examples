using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsHtmlScriptRemoval
{
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Load the HTML file into a workbook using HtmlLoadOptions
            HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
            // Optional: enable support for <div> tags if needed
            loadOptions.SupportDivTag = true;
            // Load the workbook
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Iterate through all used cells in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    // Process only cells that contain an HTML string
                    if (!string.IsNullOrEmpty(cell.HtmlString))
                    {
                        // Remove any <script>...</script> blocks (case‑insensitive, single‑line)
                        string cleaned = Regex.Replace(
                            cell.HtmlString,
                            "<script.*?</script>",
                            string.Empty,
                            RegexOptions.Singleline | RegexOptions.IgnoreCase);

                        // Update the cell with the cleaned HTML
                        cell.HtmlString = cleaned;
                    }
                }
            }

            // Save the cleaned workbook to an Excel file
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}