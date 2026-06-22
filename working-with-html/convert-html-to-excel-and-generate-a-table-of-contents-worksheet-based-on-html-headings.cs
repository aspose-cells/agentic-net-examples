using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Utility;

class HtmlToExcelWithToc
{
    static void Main()
    {
        // Paths for input HTML and output Excel files
        string htmlPath = "input.html";
        string excelPath = "output.xlsx";

        // Load the HTML file into a workbook using HtmlLoadOptions
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        Workbook workbook = new Workbook(htmlPath, loadOptions); // load

        // Read the raw HTML content for heading extraction
        string htmlContent = File.ReadAllText(htmlPath);

        // Regex to capture heading tags (h1 to h6) and their inner text
        Regex headingRegex = new Regex(@"<(h[1-6])[^>]*>(.*?)</\1>", RegexOptions.IgnoreCase);
        MatchCollection matches = headingRegex.Matches(htmlContent);

        // Create a new worksheet at the beginning to serve as Table of Contents
        Worksheet tocSheet = workbook.Worksheets[workbook.Worksheets.Add()];
        tocSheet.Name = "Table of Contents";

        // Write header for the TOC sheet
        tocSheet.Cells["A1"].PutValue("Table of Contents");
        tocSheet.Cells["A1"].GetStyle().Font.IsBold = true;

        // Populate TOC with extracted headings
        int rowIndex = 2; // start from row 2
        foreach (Match match in matches)
        {
            // Determine heading level (1-6)
            string tag = match.Groups[1].Value.ToLower(); // e.g., h2
            int level = int.Parse(tag.Substring(1));

            // Indent based on heading level
            string indent = new string(' ', (level - 1) * 4);
            string headingText = match.Groups[2].Value.Trim();

            // Write the heading text with indentation
            tocSheet.Cells[rowIndex, 0].PutValue(indent + headingText);
            rowIndex++;
        }

        // Save the workbook as an Excel file
        workbook.Save(excelPath); // save
    }
}