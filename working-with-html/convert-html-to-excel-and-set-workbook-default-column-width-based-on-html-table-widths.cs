// Title: Convert an HTML table to Excel and set workbook default column width from HTML <col> or <td> width attributes using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads an HTML file, loads it into an Aspose.Cells Workbook with HtmlLoadOptions, extracts column width values from <col> or <td> tags, computes the average pixel width, converts it to Excel column width using the (pixel‑5)/7+1 formula, assigns the result to Worksheet.Cells.StandardWidth, and saves the workbook as .xlsx. | Write a C# example that uses regular expressions to parse HTML table column widths, calculates the corresponding Excel column width, sets the worksheet's default column width, and exports the result with Aspose.Cells.
// Common Searches: how to keep HTML table column widths when converting to Excel with Aspose.Cells C# | Aspose.Cells set default column width from HTML col width attribute | C# convert HTML to .xlsx and preserve column sizing | calculate Excel column width from pixel values using Aspose.Cells | load HTML into workbook and adjust worksheet column width programmatically
// Tags: Aspose.Cells HTML to Excel conversion | worksheet default column width setting | extract HTML table column widths C# | pixel to Excel column width formula | HtmlLoadOptions workbook loading

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The sample reads an HTML file, loads it into an Aspose.Cells Workbook via HtmlLoadOptions, uses regular expressions to collect column width values from <col> or <td> tags, computes the average pixel width, converts that value to an Excel column width with the (pixel‑5)/7+1 formula, applies the calculated width to the worksheet's StandardWidth, and saves the workbook as output.xlsx, handling missing files and creating output directories as needed.
class HtmlToExcelConverter
{
    static void Main()
    {
        try
        {
            // Load HTML content from file (ensure the file exists)
            string htmlPath = "input.html";
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException($"Input HTML file not found: {htmlPath}");

            string html = File.ReadAllText(htmlPath);

            // Load the HTML into a workbook using a memory stream and HtmlLoadOptions
            Workbook workbook;
            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(html)))
            {
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                workbook = new Workbook(ms, loadOptions);
            }

            // Parse the HTML to extract column width information using regular expressions
            double totalWidth = 0;
            int widthCount = 0;

            // First try to read <col width="..."> definitions
            var colMatches = Regex.Matches(html, @"<col[^>]*\bwidth\s*=\s*[""']?(\d+)", RegexOptions.IgnoreCase);
            if (colMatches.Count > 0)
            {
                foreach (Match m in colMatches)
                {
                    if (double.TryParse(m.Groups[1].Value, out double widthPx))
                    {
                        totalWidth += widthPx;
                        widthCount++;
                    }
                }
            }
            else
            {
                // Fallback: read width attributes from the first row's <td> elements
                var tdMatches = Regex.Matches(html, @"<tr[^>]*>.*?<td[^>]*\bwidth\s*=\s*[""']?(\d+)", RegexOptions.Singleline | RegexOptions.IgnoreCase);
                foreach (Match m in tdMatches)
                {
                    if (double.TryParse(m.Groups[1].Value, out double widthPx))
                    {
                        totalWidth += widthPx;
                        widthCount++;
                    }
                }
            }

            // If we obtained any width values, compute an average and set the default column width
            if (widthCount > 0)
            {
                double avgPixelWidth = totalWidth / widthCount;

                // Approximate conversion from pixel width to Excel column width
                // Excel column width = (pixel - 5) / 7 + 1 (per Aspose.Cells documentation)
                double excelColumnWidth = (avgPixelWidth - 5) / 7 + 1;

                // Apply the calculated default column width to the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells.StandardWidth = excelColumnWidth; // Sets default column width
            }

            // Save the workbook to an Excel file (ensure the directory exists)
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
