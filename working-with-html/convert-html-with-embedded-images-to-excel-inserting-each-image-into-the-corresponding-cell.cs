// Title: Convert HTML Table with Base64 Images to Excel Cells using Aspose.Cells (C#)
// Description: Reads an HTML file containing a table, extracts <img> tags with base64 data URIs, converts the data to byte streams, and inserts each picture into the matching worksheet cell via Workbook.Pictures.Add before saving as .xlsx.
// Keywords: Aspose.Cells C# HTML to Excel | base64 image to Excel cell | insert picture Aspose.Cells | HTML table conversion .NET | extract img src data URI | Workbook.Pictures.Add example
// Common Searches: Aspose.Cells convert HTML to Excel with images | C# extract base64 images from HTML | Add picture to specific Excel cell using Aspose | Parse HTML table and export to .xlsx | How to embed HTML images into Excel cells
// Developer Intent: Place images from an HTML table into the corresponding Excel cells using Aspose.Cells.
// Use Cases: Generate a product catalog where each product photo embedded in an HTML table appears in the same cell of an Excel sheet. | Automate conversion of HTML email newsletters containing inline images into a structured Excel report with pictures aligned to their original positions. | Create inventory spreadsheets from web pages that use data‑URI images, preserving the visual layout when exported to Excel.
// AI Prompts: Suggest a robust HTML parsing method (e.g., HtmlAgilityPack) that handles multiple images per cell and supports both base64 and external URLs with Aspose.Cells. | Show how to automatically resize each inserted picture to fit its target cell while keeping the aspect ratio intact. | Provide error‑handling code that logs cells with missing or malformed base64 data and continues processing the rest of the table.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace HtmlToExcelWithImages
{
    // Reads an HTML file containing a table, extracts <img> tags with base64 data URIs, converts the data to byte streams, and inserts each picture into the matching worksheet cell via Workbook.Pictures.Add before saving as .xlsx.
    class Program
    {
        static void Main()
        {
            // Paths – adjust as needed
            string htmlPath = "input.html";
            string excelPath = "output.xlsx";

            // Load the HTML content
            string htmlContent = File.ReadAllText(htmlPath);

            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Parse the HTML table structure.
            // This simple parser assumes the HTML contains a single <table> with <tr> rows and <td> cells.
            int currentRow = 0;
            foreach (Match trMatch in Regex.Matches(htmlContent, @"<tr[^>]*>(.*?)</tr>", RegexOptions.Singleline | RegexOptions.IgnoreCase))
            {
                int currentCol = 0;
                string trInner = trMatch.Groups[1].Value;

                foreach (Match tdMatch in Regex.Matches(trInner, @"<td[^>]*>(.*?)</td>", RegexOptions.Singleline | RegexOptions.IgnoreCase))
                {
                    string tdInner = tdMatch.Groups[1].Value;

                    // Look for an <img> tag inside the cell
                    Match imgMatch = Regex.Match(tdInner, @"<img[^>]*src=['""]([^'""]+)['""][^>]*>", RegexOptions.IgnoreCase);
                    if (imgMatch.Success)
                    {
                        string src = imgMatch.Groups[1].Value.Trim();

                        // Handle Base64‑encoded images (data URI scheme)
                        if (src.StartsWith("data:image", StringComparison.OrdinalIgnoreCase))
                        {
                            int base64Pos = src.IndexOf("base64,", StringComparison.OrdinalIgnoreCase);
                            if (base64Pos > -1)
                            {
                                string base64Data = src.Substring(base64Pos + 7);
                                byte[] imageBytes = Convert.FromBase64String(base64Data);

                                // Insert the image into the corresponding cell (lifecycle rule: add picture)
                                using (MemoryStream ms = new MemoryStream(imageBytes))
                                {
                                    sheet.Pictures.Add(currentRow, currentCol, ms);
                                }
                            }
                        }
                        // Optional: handle external image files or URLs (not required for the core task)
                    }

                    currentCol++;
                }

                currentRow++;
            }

            // Save the workbook to an Excel file (lifecycle rule: save)
            workbook.Save(excelPath);
            Console.WriteLine($"HTML converted to Excel with images saved at: {excelPath}");
        }
    }
}
