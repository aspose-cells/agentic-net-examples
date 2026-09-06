// Title: Convert an HTML string with inline SVG to an Excel .xlsx file while keeping SVG graphics as scalable vector shapes using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an HTML string containing <svg> elements into an Aspose.Cells Workbook with HtmlLoadOptions and inserts each SVG as a vector picture using Worksheets.Pictures.Add. | Show how to extract SVG markup from the HTML using Regex, wrap it in a MemoryStream, place each SVG at successive rows on a worksheet, and save the workbook to a specified folder while creating the folder if it does not exist.
// Common Searches: how to keep SVG images vector when converting HTML to Excel with Aspose.Cells | Aspose.Cells C# load HTML string and embed inline SVG as pictures | convert HTML containing <svg> tags to .xlsx preserving vector quality | add multiple SVG graphics from HTML to Excel worksheet using Aspose.Cells | save Excel file to a new folder after converting HTML with SVG in C#
// Tags: html to xlsx conversion with svg preservation | aspocells add svg picture from stream | load html into workbook using HtmlLoadOptions | c# regex extract inline svg | worksheet pictures add vector graphics

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example reads an HTML string that includes inline <svg> elements, loads it into a Workbook via HtmlLoadOptions, uses a regular expression to locate each SVG, converts the markup to a MemoryStream, adds the SVGs as vector pictures to the first worksheet with Pictures.Add, and saves the workbook as an .xlsx file, creating the output directory if necessary.
class HtmlToExcelWithSvg
{
    static void Main()
    {
        try
        {
            // Sample HTML input containing SVG graphics
            string htmlContent = @"
                <html>
                    <body>
                        <h1>Report</h1>
                        <p>Data overview:</p>
                        <svg width='100' height='100' xmlns='http://www.w3.org/2000/svg'>
                            <circle cx='50' cy='50' r='40' stroke='green' stroke-width='4' fill='yellow' />
                        </svg>
                        <p>End of report.</p>
                    </body>
                </html>";

            // Convert HTML string to a memory stream
            byte[] htmlBytes = Encoding.UTF8.GetBytes(htmlContent);
            using (var htmlStream = new MemoryStream(htmlBytes))
            {
                // Load HTML into a new workbook using HtmlLoadOptions
                var loadOptions = new HtmlLoadOptions();
                var workbook = new Workbook(htmlStream, loadOptions);

                // Extract SVG elements and add them as pictures
                var worksheet = workbook.Worksheets[0];
                var svgMatches = Regex.Matches(htmlContent, @"<svg[\s\S]*?<\/svg>", RegexOptions.IgnoreCase);
                if (svgMatches.Count > 0)
                {
                    int startRow = 0;
                    foreach (Match match in svgMatches)
                    {
                        string svgMarkup = match.Value;
                        byte[] svgBytes = Encoding.UTF8.GetBytes(svgMarkup);
                        using (var svgStream = new MemoryStream(svgBytes))
                        {
                            // Add SVG picture at the specified cell
                            worksheet.Pictures.Add(startRow, 0, svgStream);
                        }
                        startRow += 15; // offset for next SVG
                    }
                }

                // Save workbook
                string outputPath = "OutputWithSvg.xlsx";
                try
                {
                    // Ensure the directory exists
                    string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Error saving workbook: {saveEx.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
