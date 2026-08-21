// Title: C# – Convert HTML to Excel with Aspose.Cells and keep the body background image as worksheet background
// Description: A complete C# example that loads an HTML file into an Aspose.Cells Workbook, extracts the <body> background image (from a background attribute or CSS background‑image style), resolves relative paths, inserts the image as a picture on each worksheet, and saves the workbook as an XLSX file.
// Keywords: Aspose.Cells | HTML to Excel conversion | C# | .NET | worksheet background image | preserve HTML background | LoadOptions Html | SaveFormat Xlsx | regular expression image extraction | batch HTML to XLSX
// Common Searches: Aspose.Cells keep HTML body background when converting to XLSX | C# add worksheet background picture after loading HTML | extract background-image URL from HTML for Excel workbook | convert HTML to Excel with background image using Aspose.Cells | load HTML with background attribute in Aspose.Cells .NET
// Developer Intent: Insert the HTML page’s background image into every worksheet of the generated Excel file.
// Use Cases: Create branded Excel reports that retain a logo or watermark defined as a page background in an HTML template. | Generate printable spreadsheets that visually match a web form by preserving its background image during conversion. | Automate batch conversion of multiple HTML files to Excel while automatically applying detected background images to each worksheet.
// AI Prompts: Write C# code with Aspose.Cells to load an HTML file, detect the <body> background image (attribute or CSS), and add it as a worksheet background before saving as XLSX. | Explain how to resolve relative image paths when converting HTML to Excel using Aspose.Cells so the picture appears on all worksheets. | Provide a step‑by‑step guide for extracting a background‑image URL from HTML using regular expressions and applying it to a workbook with Aspose.Cells.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// A complete C# example that loads an HTML file into an Aspose.Cells Workbook, extracts the <body> background image (from a background attribute or CSS background‑image style), resolves relative paths, inserts the image as a picture on each worksheet, and saves the workbook as an XLSX file.
class HtmlToExcelWithBackground
{
    static void Main()
    {
        // Paths for input HTML and output Excel files
        string htmlFilePath = "input.html";
        string excelFilePath = "output.xlsx";

        try
        {
            // Verify that the HTML file exists
            if (!File.Exists(htmlFilePath))
            {
                Console.WriteLine($"HTML file not found: {htmlFilePath}");
                return;
            }

            // Load the HTML file into a workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(htmlFilePath, loadOptions);

            // ------------------------------------------------------------
            // Extract background image URL from the HTML file (simple approach)
            // Supports <body background="image.jpg"> or CSS style:
            //   <body style="background-image:url('image.jpg')">
            // ------------------------------------------------------------
            string htmlContent = File.ReadAllText(htmlFilePath);
            string bgImagePath = null;

            // Try <body background="...">
            Match match = Regex.Match(
                htmlContent,
                @"<body[^>]*\sbackground\s*=\s*[""']([^""']+)[""']",
                RegexOptions.IgnoreCase);
            if (match.Success)
            {
                bgImagePath = match.Groups[1].Value;
            }
            else
            {
                // Try CSS background-image in style attribute
                match = Regex.Match(
                    htmlContent,
                    @"<body[^>]*\sstyle\s*=\s*[""'][^""']*background-image\s*:\s*url\(['""]?([^'"")]+)['""]?\)[^""']*[""']",
                    RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    bgImagePath = match.Groups[1].Value;
                }
            }

            // If a background image was found, insert it into each worksheet
            if (!string.IsNullOrEmpty(bgImagePath))
            {
                // Resolve relative paths based on the HTML file location
                string resolvedPath = Path.IsPathRooted(bgImagePath)
                    ? bgImagePath
                    : Path.Combine(Path.GetDirectoryName(htmlFilePath) ?? string.Empty, bgImagePath);

                if (File.Exists(resolvedPath))
                {
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Insert the image as a picture covering the sheet (top‑left cell 0,0)
                        sheet.Pictures.Add(0, 0, resolvedPath);
                    }
                }
                else
                {
                    Console.WriteLine($"Background image file not found: {resolvedPath}");
                }
            }
            else
            {
                Console.WriteLine("No background image detected in the HTML file.");
            }

            // Save the workbook as an Excel file
            workbook.Save(excelFilePath, SaveFormat.Xlsx);
            Console.WriteLine($"Conversion completed. Excel saved to: {excelFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
