// Title: How to remove <script> tags from HTML cells after loading an HTML file into an Aspose.Cells workbook using C#
// AI Prompts: Write C# code that loads an HTML document with HtmlLoadOptions, iterates through every worksheet and cell, strips <script> elements from each cell's HtmlString using a regular expression, and saves the workbook as XLSX. | Create a reusable C# method that accepts a Workbook object, cleans all HtmlString values by removing script tags with Regex, and returns the sanitized workbook ready for export.
// Common Searches: C# Aspose.Cells remove script tags from imported HTML | How to strip JavaScript from cells after loading HTML in Aspose.Cells | Cleaning HtmlString in Excel cells using Aspose.Cells C# | Prevent script execution when converting HTML to XLSX with Aspose.Cells | Regex to delete <script> elements from workbook cells in C#
// Tags: Aspose.Cells HTMLLoadOptions script tag removal | C# regex clean HtmlString in workbook cells | sanitize cell content after HTML import Aspose.Cells | remove embedded JavaScript from Excel cells C# | Aspose.Cells export HTML to XLSX without scripts

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsHtmlScriptRemoval
{
    // The program loads an HTML file into an Aspose.Cells workbook, walks through each worksheet and cell, removes any <script> tags from the cell's HtmlString using a regular expression, and saves the cleaned workbook as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Load the HTML file into a workbook using HtmlLoadOptions
            HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook("input.html", loadOptions);

            // Regular expression to match <script> tags (including their content)
            string scriptPattern = @"<script\b[^>]*?>.*?</script>";
            Regex scriptRegex = new Regex(scriptPattern, RegexOptions.Singleline | RegexOptions.IgnoreCase);

            // Iterate through all worksheets and cells to clean script tags
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                // Get the maximum used row and column to limit the iteration
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        // Process only if the cell contains an HTML string
                        if (!string.IsNullOrEmpty(cell.HtmlString))
                        {
                            // Remove all <script> tags from the HTML content
                            string cleanedHtml = scriptRegex.Replace(cell.HtmlString, string.Empty);
                            // Update the cell with the cleaned HTML
                            cell.HtmlString = cleanedHtml;
                        }
                    }
                }
            }

            // Save the cleaned workbook to an Excel file
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
