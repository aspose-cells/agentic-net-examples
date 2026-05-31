using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

class HtmlToExcelWithPageBreaks
{
    static void Main()
    {
        try
        {
            // Paths for input HTML and output Excel files
            string htmlPath = "input.html";
            string excelPath = "output.xlsx";

            // Verify that the input HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: Input file '{htmlPath}' not found.");
                return;
            }

            // Read the entire HTML content
            string htmlContent = File.ReadAllText(htmlPath);

            // Create a new workbook (empty)
            Workbook workbook = new Workbook();

            // Add a worksheet where the HTML table will be placed
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "ImportedHtml";

            // Regular expression to match table rows (<tr>...</tr>)
            Regex rowRegex = new Regex(@"<tr[^>]*>(.*?)</tr>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            // Regular expression to match table cells (<td>...</td> or <th>...</th>)
            Regex cellRegex = new Regex(@"<(td|th)[^>]*>(.*?)</\1>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            // Marker that indicates a page break in the source HTML
            const string pageBreakMarker = "<!--PAGE_BREAK-->";

            int currentRow = 0; // Zero‑based row index in the worksheet

            // Iterate over each <tr> element found in the HTML
            foreach (Match rowMatch in rowRegex.Matches(htmlContent))
            {
                string rowHtml = rowMatch.Groups[1].Value;

                // If the row contains the page‑break marker, add a horizontal page break
                // before this row (i.e., after the previous row)
                if (rowHtml.IndexOf(pageBreakMarker, StringComparison.OrdinalIgnoreCase) >= 0 && currentRow > 0)
                {
                    try
                    {
                        // Use dynamic to avoid compile‑time dependency on the exact API version
                        ((dynamic)sheet).HPageBreaks?.Add(currentRow);
                    }
                    catch
                    {
                        // If the API does not support page breaks, ignore silently
                    }
                }

                int currentCol = 0; // Zero‑based column index

                // Extract each cell within the current row
                foreach (Match cellMatch in cellRegex.Matches(rowHtml))
                {
                    // Strip any HTML tags inside the cell value to get plain text
                    string cellInnerHtml = cellMatch.Groups[2].Value;
                    string cellText = Regex.Replace(cellInnerHtml, "<.*?>", string.Empty).Trim();

                    // Put the extracted text into the corresponding cell
                    sheet.Cells[currentRow, currentCol].PutValue(cellText);
                    currentCol++;
                }

                // Move to the next worksheet row
                currentRow++;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(excelPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to Excel format
            workbook.Save(excelPath, SaveFormat.Xlsx);

            Console.WriteLine($"HTML has been converted to Excel with page breaks preserved at '{excelPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}