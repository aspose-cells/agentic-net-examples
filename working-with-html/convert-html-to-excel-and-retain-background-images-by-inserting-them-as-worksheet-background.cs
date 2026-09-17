// Title: Convert HTML to an Excel workbook and preserve CSS background images as worksheet backgrounds using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that reads an HTML file, parses the first CSS background‑image URL, loads the HTML into an Aspose.Cells Workbook with HtmlLoadOptions, inserts the image as a picture on every worksheet, and saves the workbook as an XLSX file. | Show how to programmatically set a worksheet background picture in Aspose.Cells after converting HTML to a workbook, handling both absolute and relative image paths.
// Common Searches: how to keep background image when converting html to excel with Aspose.Cells C# | c# parse CSS background-image URL and apply as worksheet background in Aspose.Cells | Aspose.Cells HtmlLoadOptions preserve CSS background image in generated XLSX | add same picture to all worksheets after loading html file in Aspose.Cells | convert html file to xlsx and set worksheet background picture using .NET
// Tags: Aspose.Cells HTML to XLSX conversion with background image | C# parse CSS for background-image URL | apply background picture to worksheets Aspose.Cells | use HtmlLoadOptions for HTML import Aspose.Cells | insert image into all worksheets via Pictures.Add

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example reads an HTML file, extracts the first CSS background‑image URL (handling relative paths), loads the HTML into an Aspose.Cells Workbook using HtmlLoadOptions, adds the image as a picture anchored at cell A1 on each worksheet, and saves the result as an XLSX workbook.
class HtmlToExcelWithBackground
{
    static void Main()
    {
        try
        {
            // Path to the source HTML file
            string htmlPath = @"C:\Input\sample.html";

            // Path where the resulting Excel file will be saved
            string excelPath = @"C:\Output\result.xlsx";

            // Verify that the HTML source file exists
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException("HTML source file not found.", htmlPath);

            // --------------------------------------------------------------------
            // 1. Load the HTML content into a string
            // --------------------------------------------------------------------
            string htmlContent = File.ReadAllText(htmlPath);

            // --------------------------------------------------------------------
            // 2. Extract the first background-image URL from the HTML (if any)
            // --------------------------------------------------------------------
            string bgImagePath = null;
            var match = Regex.Match(
                htmlContent,
                @"background-image\s*:\s*url\(['""]?(?<url>[^'"")]+)['""]?\)",
                RegexOptions.IgnoreCase);
            if (match.Success)
            {
                bgImagePath = match.Groups["url"].Value;
                // If the URL is relative, combine it with the HTML file directory
                if (!Path.IsPathRooted(bgImagePath))
                {
                    string htmlDir = Path.GetDirectoryName(htmlPath) ?? string.Empty;
                    bgImagePath = Path.GetFullPath(Path.Combine(htmlDir, bgImagePath));
                }
            }

            // --------------------------------------------------------------------
            // 3. Load the HTML into a new Workbook instance
            // --------------------------------------------------------------------
            // Aspose.Cells loads HTML from a file; write the HTML string to a temp file
            string tempHtmlPath = Path.GetTempFileName();
            File.WriteAllText(tempHtmlPath, htmlContent);
            var loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(tempHtmlPath, loadOptions);
            File.Delete(tempHtmlPath); // clean up temporary file

            // --------------------------------------------------------------------
            // 4. If a background image was found, add it as a picture to each worksheet
            // --------------------------------------------------------------------
            if (!string.IsNullOrEmpty(bgImagePath) && File.Exists(bgImagePath))
            {
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Add the image as a picture anchored at the top‑left corner (cell A1)
                    sheet.Pictures.Add(0, 0, bgImagePath);
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(excelPath) ?? string.Empty;
            if (!Directory.Exists(outputDir) && !string.IsNullOrEmpty(outputDir))
                Directory.CreateDirectory(outputDir);

            // --------------------------------------------------------------------
            // 5. Save the workbook to an Excel file
            // --------------------------------------------------------------------
            workbook.Save(excelPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
