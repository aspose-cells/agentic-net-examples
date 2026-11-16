// Title: Convert HTML to Excel while preserving HTML page‑break tags as worksheet page breaks using Aspose.Cells for .NET (C#)
// Common Searches: aspnet convert html file to xlsx and retain page-break-before elements | c# Aspose.Cells replace HTML page-break-after with Excel horizontal page break | how to map HTML page break markers to Excel worksheet page breaks in C# | load modified HTML into Aspose.Cells workbook to keep original pagination | regex replace page-break style in HTML before importing to Aspose.Cells
// Tags: Aspose.Cells HTML to XLSX conversion with page break mapping | C# regex placeholder for HTML page-break handling | Insert horizontal page breaks from cell placeholder Aspose.Cells | Workbook load from temporary HTML file Aspose.Cells | Clear placeholder after adding worksheet page break | Save workbook as Xlsx using Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example reads an HTML document, uses a regular expression to replace elements styled with page-break-before or page-break-after: always with a unique placeholder, loads the modified HTML into an Aspose.Cells Workbook, scans the cells for the placeholder, inserts a horizontal page break before the corresponding row, removes the placeholder text, and saves the result as an XLSX file.
class HtmlToExcelWithPageBreaks
{
    static void Main()
    {
        // Paths for input HTML and output Excel files
        string htmlPath = @"C:\Input\sample.html";
        string excelPath = @"C:\Output\result.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException($"Input HTML file not found: {htmlPath}");

            // Ensure output directory exists
            string? outputDir = Path.GetDirectoryName(excelPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Read the HTML content
            string htmlContent = File.ReadAllText(htmlPath);

            // Define a unique placeholder that will represent a page break in the worksheet
            const string pageBreakPlaceholder = "[PAGE_BREAK_MARKER]";

            // Replace HTML page‑break markers with the placeholder.
            string pattern = @"<(?<tag>\w+)([^>]*?)style\s*=\s*[""'][^""']*page-break-(before|after)\s*:\s*always[^""']*[""'][^>]*?>\s*</\k<tag>>";
            htmlContent = Regex.Replace(htmlContent, pattern, $"<span>{pageBreakPlaceholder}</span>", RegexOptions.IgnoreCase);

            // Write modified HTML to a temporary file for loading
            string tempHtmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".html");
            File.WriteAllText(tempHtmlPath, htmlContent);

            // Load the modified HTML into an Aspose.Cells workbook
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(tempHtmlPath, loadOptions);

            // Delete temporary file
            File.Delete(tempHtmlPath);

            // Get the first worksheet (adjust if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Scan all cells to locate the placeholder and insert a page break before that row
            Cells cells = sheet.Cells;
            for (int row = 0; row <= cells.MaxDataRow; row++)
            {
                for (int col = 0; col <= cells.MaxDataColumn; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.Value != null && cell.StringValue.Contains(pageBreakPlaceholder))
                    {
                        // Insert a horizontal page break before this row
                        sheet.HorizontalPageBreaks.Add(row);

                        // Clear the placeholder text from the cell
                        cell.PutValue(string.Empty);
                    }
                }
            }

            // Save the workbook to the desired Excel format
            workbook.Save(excelPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
