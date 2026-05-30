using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace HtmlToExcelAlignment
{
    class Program
    {
        static void Main()
        {
            // Paths for input HTML and output Excel files
            string htmlPath = "input.html";
            string excelPath = "output.xlsx";

            // Ensure the HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"HTML file not found: {htmlPath}");
                return;
            }

            // Read the entire HTML content
            string htmlContent = File.ReadAllText(htmlPath);

            // Create a new workbook (empty Excel file)
            Workbook workbook = new Workbook();

            // Get the first worksheet and its cells collection
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Regular expression to capture each <tr> block
            Regex rowRegex = new Regex(@"<tr[^>]*>(.*?)</tr>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            // Regular expression to capture each <td> with a text-align style
            Regex cellRegex = new Regex(
                @"<td[^>]*style\s*=\s*""[^""]*text-align\s*:\s*(left|center|right)[^""]*""[^>]*>(.*?)</td>",
                RegexOptions.IgnoreCase);

            // Iterate over rows
            int rowIndex = 0;
            foreach (Match rowMatch in rowRegex.Matches(htmlContent))
            {
                string rowInnerHtml = rowMatch.Groups[1].Value;

                // Iterate over cells within the current row
                int colIndex = 0;
                foreach (Match cellMatch in cellRegex.Matches(rowInnerHtml))
                {
                    // Extract alignment value (left, center, right)
                    string alignment = cellMatch.Groups[1].Value.Trim().ToLower();

                    // Extract cell text (strip any inner HTML tags)
                    string rawCellText = cellMatch.Groups[2].Value;
                    string cellText = Regex.Replace(rawCellText, "<.*?>", string.Empty).Trim();

                    // Put the text into the corresponding Excel cell
                    Cell cell = cells[rowIndex, colIndex];
                    cell.PutValue(cellText);

                    // Get the current style, modify horizontal alignment, and apply it back
                    Style style = cell.GetStyle();
                    switch (alignment)
                    {
                        case "left":
                            style.HorizontalAlignment = TextAlignmentType.Left;
                            break;
                        case "center":
                            style.HorizontalAlignment = TextAlignmentType.Center;
                            break;
                        case "right":
                            style.HorizontalAlignment = TextAlignmentType.Right;
                            break;
                    }
                    cell.SetStyle(style);

                    colIndex++;
                }

                rowIndex++;
            }

            // Save the workbook to an Excel file using default Ooxml options
            workbook.Save(excelPath, SaveFormat.Xlsx);

            Console.WriteLine($"HTML has been converted to Excel with alignment mapping: {excelPath}");
        }
    }
}