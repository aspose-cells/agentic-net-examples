// Title: Convert an HTML file to Excel (.xlsx) while preserving HTML page‑break elements as worksheet page breaks using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a local HTML document into an Aspose.Cells Workbook, scans the markup for <hr> tags and CSS page‑break directives, and inserts matching horizontal page breaks into the first worksheet. | Show how to map block‑level HTML elements to Excel rows, track the current row index, and add page breaks at the appropriate rows before saving the workbook. | Provide a C# snippet that creates the output directory if it does not exist, saves the workbook as an .xlsx file, and includes error handling for the conversion process.
// Common Searches: how to keep HTML <hr> page breaks when converting to XLSX with Aspose.Cells in C# | Aspose.Cells C# convert HTML to Excel preserving CSS page-break-before | add worksheet page breaks from HTML markers using Aspose.Cells | detect CSS page-break-after in HTML and apply Excel pagination | C# load HTML into Aspose.Cells workbook and customize pagination
// Tags: Aspose.Cells HTML to XLSX conversion with pagination | map HTML block elements to Excel rows C# | identify HTML page-break markers Aspose.Cells | load HTML using HtmlLoadOptions Aspose.Cells | save workbook as .xlsx with custom page breaks

using System;
using System.IO;
using Aspose.Cells;

// The example reads an HTML file, loads it into an Aspose.Cells Workbook, scans the markup for <hr> tags and CSS page‑break directives, adds matching horizontal page breaks to the first worksheet, creates the output folder if needed, and saves the result as an .xlsx file.
class HtmlToExcelWithPageBreaks
{
    static void Main()
    {
        try
        {
            // Path to the source HTML file
            string htmlPath = @"C:\Input\source.html";

            // Path to the output Excel file
            string excelPath = @"C:\Output\result.xlsx";

            // Verify that the input HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Input file not found: {htmlPath}");
                return;
            }

            // Load the HTML content into a string (needed for page‑break detection)
            string htmlContent = File.ReadAllText(htmlPath);

            // -------------------------------------------------
            // 1. Load the HTML into a new Workbook instance
            // -------------------------------------------------
            Workbook workbook;
            try
            {
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                workbook = new Workbook(htmlPath, loadOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load HTML into workbook: {ex.Message}");
                return;
            }

            // -------------------------------------------------
            // 2. Detect HTML page‑break markers and add worksheet page breaks
            // -------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];

            // Approximate current Excel row while scanning the HTML content line by line
            int currentRow = 0;

            // Helper to add a horizontal page break after the specified row
            void AddPageBreak(int rowIndex)
            {
                // Ensure the row index is within a reasonable range
                if (rowIndex >= 0 && rowIndex <= sheet.Cells.MaxDataRow + 1)
                {
                    // Use HorizontalPageBreaks collection (compatible with various Aspose.Cells versions)
                    sheet.HorizontalPageBreaks.Add(rowIndex);
                }
            }

            // Scan the HTML content line by line
            using (StringReader reader = new StringReader(htmlContent))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    string lower = line.ToLowerInvariant();

                    // Detect page‑break markers
                    bool isPageBreak = lower.Contains("<hr") ||
                                       lower.Contains("page-break-before:always") ||
                                       lower.Contains("page-break-after:always");

                    if (isPageBreak)
                    {
                        // Add a page break after the current row
                        AddPageBreak(currentRow);
                    }

                    // Approximate row advancement:
                    // Treat block‑level tags as generating a new row.
                    if (lower.Contains("<p") ||
                        lower.Contains("<div") ||
                        lower.Contains("<tr") ||
                        lower.Contains("<li") ||
                        lower.Contains("<h1") || lower.Contains("<h2") ||
                        lower.Contains("<h3") || lower.Contains("<h4") ||
                        lower.Contains("<h5") || lower.Contains("<h6"))
                    {
                        currentRow++;
                    }
                    else
                    {
                        // Count line breaks within the line's text (if any)
                        int lineBreaks = line.Split('\n').Length - 1;
                        currentRow += lineBreaks;
                    }
                }
            }

            // -------------------------------------------------
            // 3. Save the workbook to an Excel file
            // -------------------------------------------------
            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(excelPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(excelPath, SaveFormat.Xlsx);

            Console.WriteLine("HTML has been converted to Excel with page breaks preserved.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
