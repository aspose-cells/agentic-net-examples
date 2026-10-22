// Title: Convert HTML to Excel with Aspose.Cells in C# while preserving the body background image as worksheet background
// AI Prompts: Write C# code that uses Aspose.Cells to load an HTML file into a Workbook, parses the HTML or embedded CSS to locate the body’s background‑image reference, adds that image as a picture to every worksheet (simulating a background), and saves the result as an XLSX file. | Create a reusable C# method that accepts the path of an HTML document, resolves any relative background‑image location, inserts the image into a worksheet using Aspose.Cells Pictures.Add, returns the modified Workbook, and gracefully handles missing image files.
// Common Searches: how to keep background image when converting html to xlsx using aspose.cells c# | asp.net extract body background-image from html and apply to excel worksheet | c# load html into workbook with aspose.cells and set worksheet background picture | convert html page to excel preserving css background image in .net
// Tags: aspose.cells html to xlsx conversion with worksheet background picture | c# extract css background-image reference from html | aspose.cells add picture to worksheet | handle relative image locations for worksheet background | aspose.cells import html using LoadOptions Html

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// // This program reads an HTML file, loads it into an Aspose.Cells Workbook, extracts any body background-image reference from inline or CSS styles, adds the image as a picture to each worksheet to act as a background, and saves the workbook as an XLSX file.
class HtmlToExcelWithBackground
{
    static void Main()
    {
        // Paths for input HTML and output Excel
        string htmlPath = @"C:\Input\sample.html";
        string excelPath = @"C:\Output\result.xlsx";

        try
        {
            // Verify that the HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"HTML file not found: {htmlPath}");
                return;
            }

            // Load HTML content
            string htmlContent = File.ReadAllText(htmlPath);

            // -------------------------------------------------
            // 1. Load HTML into a new Workbook
            // -------------------------------------------------
            // Use the Workbook constructor that accepts a stream and LoadOptions
            using (MemoryStream ms = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)))
            {
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                Workbook workbook = new Workbook(ms, loadOptions);

                // -------------------------------------------------
                // 2. Extract background image URL from the HTML
                // -------------------------------------------------
                string bgImageUrl = null;
                var bodyStyleMatch = Regex.Match(
                    htmlContent,
                    @"<body[^>]*style\s*=\s*[""']([^""']*background-image\s*:\s*url\(['""]?(?<url>[^)'""]+)['""]?\))[^>]*>",
                    RegexOptions.IgnoreCase);
                if (bodyStyleMatch.Success)
                {
                    bgImageUrl = bodyStyleMatch.Groups["url"].Value;
                }
                else
                {
                    var styleMatch = Regex.Match(
                        htmlContent,
                        @"body\s*{[^}]*background-image\s*:\s*url\(['""]?(?<url>[^)'""]+)['""]?\)",
                        RegexOptions.IgnoreCase);
                    if (styleMatch.Success)
                    {
                        bgImageUrl = styleMatch.Groups["url"].Value;
                    }
                }

                // -------------------------------------------------
                // 3. If a background image was found, add it to each worksheet
                // -------------------------------------------------
                if (!string.IsNullOrEmpty(bgImageUrl))
                {
                    // Resolve relative paths based on the HTML file location
                    string imagePath = Path.IsPathRooted(bgImageUrl)
                        ? bgImageUrl
                        : Path.Combine(Path.GetDirectoryName(htmlPath) ?? string.Empty, bgImageUrl);

                    if (File.Exists(imagePath))
                    {
                        // Add the image as a picture covering the worksheet (acts as a background)
                        foreach (Worksheet sheet in workbook.Worksheets)
                        {
                            // Add picture at cell A1; you can adjust positioning as needed
                            sheet.Pictures.Add(0, 0, imagePath);
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Background image file not found: {imagePath}");
                    }
                }
                else
                {
                    Console.WriteLine("No background image detected in the HTML.");
                }

                // -------------------------------------------------
                // 4. Save the workbook to an Excel file
                // -------------------------------------------------
                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(excelPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(excelPath, SaveFormat.Xlsx);
                Console.WriteLine($"HTML has been converted to Excel and saved at: {excelPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
