// Title: Convert HTML to XLSX while preserving the page background image as a worksheet background using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells convert HTML file to Excel and keep CSS background-image as sheet background | how to extract background-image URL from HTML and set it as worksheet background in Aspose.Cells | programmatically add a free‑floating picture as worksheet background after loading HTML with Aspose.Cells | download remote background image from HTML and apply it to an Excel workbook using Aspose.Cells C# | load HTML into Aspose.Cells workbook and insert background picture on first sheet
// Tags: HTML to XLSX conversion with Aspose.Cells C# | worksheet background picture insertion Aspose.Cells | extract CSS background-image URL C# | download remote image for Excel sheet Aspose.Cells | freefloating picture placement Aspose.Cells workbook

using System;
using System.IO;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example reads an HTML file, extracts the first CSS background‑image URL using a regular expression, loads the HTML into an Aspose.Cells Workbook, resolves the image path (downloading remote images or locating local files), inserts the image as a free‑floating picture on the first worksheet to simulate a sheet background, and saves the workbook as an XLSX file.
class Program
{
    static async Task Main(string[] args)
    {
        // ------------------------------------------------------------
        // 1. Determine input HTML file and output Excel file paths
        // ------------------------------------------------------------
        string htmlPath = args.Length > 0 ? args[0] : "input.html";
        string excelPath = args.Length > 1 ? args[1] : "output.xlsx";

        if (!File.Exists(htmlPath))
        {
            Console.WriteLine($"Error: HTML file not found at '{htmlPath}'.");
            return;
        }

        // ------------------------------------------------------------
        // 2. Read HTML content
        // ------------------------------------------------------------
        string htmlContent = File.ReadAllText(htmlPath);

        // ------------------------------------------------------------
        // 3. Extract the first background‑image URL from the HTML
        // ------------------------------------------------------------
        string bgImageUrl = ExtractBackgroundImageUrl(htmlContent);
        if (string.IsNullOrEmpty(bgImageUrl))
        {
            Console.WriteLine("Warning: No background image URL found in the HTML. The workbook will be created without a sheet background.");
        }

        // ------------------------------------------------------------
        // 4. Load the HTML into an Aspose.Cells Workbook
        // ------------------------------------------------------------
        Workbook workbook;
        try
        {
            workbook = new Workbook(htmlPath); // Aspose.Cells detects HTML format automatically
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading HTML into workbook: {ex.Message}");
            return;
        }

        // ------------------------------------------------------------
        // 5. If a background image was found, add it as a picture on the first worksheet
        // ------------------------------------------------------------
        if (!string.IsNullOrEmpty(bgImageUrl))
        {
            try
            {
                string baseDir = Path.GetDirectoryName(htmlPath) ?? Directory.GetCurrentDirectory();
                string localImagePath = await ResolveImagePath(bgImageUrl, baseDir);

                if (string.IsNullOrEmpty(localImagePath) || !File.Exists(localImagePath))
                {
                    Console.WriteLine($"Error: Resolved image path is invalid: '{localImagePath}'.");
                }
                else
                {
                    Worksheet sheet = workbook.Worksheets[0];

                    // Insert the image at the top‑left corner (cell A1). This mimics a sheet background.
                    int pictureIndex = sheet.Pictures.Add(0, 0, localImagePath);
                    Picture pic = sheet.Pictures[pictureIndex];
                    pic.Placement = PlacementType.FreeFloating; // ensure it stays behind cells

                    Console.WriteLine($"Background image applied from '{localImagePath}'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error applying background image: {ex.Message}");
            }
        }

        // ------------------------------------------------------------
        // 6. Save the workbook as Excel
        // ------------------------------------------------------------
        try
        {
            workbook.Save(excelPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook successfully saved to '{excelPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }

    private static string ExtractBackgroundImageUrl(string html)
    {
        // Regex looks for url(...) inside a background or background-image property
        var regex = new Regex(@"background(?:-image)?\s*:\s*url\(\s*['""]?(?<url>[^'"")]+)['""]?\s*\)", RegexOptions.IgnoreCase);
        var match = regex.Match(html);
        if (match.Success)
        {
            return match.Groups["url"].Value.Trim();
        }

        // Fallback: look for <body ... style="...background-image:url(...)...">
        var bodyStyleRegex = new Regex(@"<body[^>]*\sstyle\s*=\s*[""'](?<style>[^""']+)[""']", RegexOptions.IgnoreCase);
        var bodyMatch = bodyStyleRegex.Match(html);
        if (bodyMatch.Success)
        {
            var style = bodyMatch.Groups["style"].Value;
            var styleMatch = regex.Match(style);
            if (styleMatch.Success)
                return styleMatch.Groups["url"].Value.Trim();
        }

        return null;
    }

    private static async Task<string> ResolveImagePath(string url, string baseDirectory)
    {
        // Remote image – download it to a temporary file
        if (Uri.TryCreate(url, UriKind.Absolute, out Uri uri) &&
            (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps))
        {
            try
            {
                using HttpClient client = new HttpClient();
                byte[] data = await client.GetByteArrayAsync(uri);
                string tempFile = Path.Combine(Path.GetTempPath(),
                    Guid.NewGuid().ToString() + Path.GetExtension(uri.AbsolutePath));
                await File.WriteAllBytesAsync(tempFile, data);
                return tempFile;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error downloading remote image: {ex.Message}");
                return null;
            }
        }

        // Assume local file path (relative to the HTML file location)
        string combinedPath = Path.IsPathRooted(url) ? url : Path.Combine(baseDirectory, url);
        if (!File.Exists(combinedPath))
        {
            Console.WriteLine($"Error: Background image file not found at '{combinedPath}'.");
            return null;
        }
        return combinedPath;
    }
}
