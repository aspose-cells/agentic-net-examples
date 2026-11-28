// Title: Export Aspose.Cells chart images as PNG when saving a workbook to HTML and verify the PNG references in C#
// Common Searches: how to export chart images as PNG using Aspose.Cells HTMLSaveOptions in C# | C# code to check that chart images in Aspose.Cells generated HTML are PNG files | Aspose.Cells save workbook to HTML with PNG chart images example | validate img src extensions for charts in HTML output from Aspose.Cells .NET | verify chart image format after exporting workbook to HTML with Aspose.Cells
// Tags: Aspose.Cells HTMLSaveOptions chart PNG export | C# chart image format verification in HTML | Aspose.Cells generated HTML image path validation | export workbook charts as PNG using Aspose.Cells | regex extract img src from Aspose.Cells HTML output

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Saving;

// The sample creates a workbook, adds a column chart, and saves it to HTML using Aspose.Cells HtmlSaveOptions, which outputs chart images as PNG files. It then reads the generated HTML, extracts all <img> src attributes with a regular expression, resolves each image path, and confirms that every referenced chart image has a .png extension and exists on disk.
class Program
{
    static void Main()
    {
        try
        {
            // Path for the output HTML file
            string htmlPath = "ChartExport.html";

            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate data for the chart
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue("A");
            cells["A3"].PutValue("B");
            cells["A4"].PutValue("C");
            cells["B2"].PutValue(10);
            cells["B3"].PutValue(20);
            cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure HTML save options (default chart image format is PNG)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as HTML
            workbook.Save(htmlPath, htmlOptions);
            Console.WriteLine($"Workbook saved as HTML to '{htmlPath}' with chart images in PNG format.");

            // Verify that the generated HTML references PNG files for chart images
            VerifyChartImagesArePng(htmlPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    static void VerifyChartImagesArePng(string htmlFilePath)
    {
        if (!File.Exists(htmlFilePath))
        {
            Console.WriteLine("HTML file not found. Verification aborted.");
            return;
        }

        // Read the HTML content
        string htmlContent = File.ReadAllText(htmlFilePath);

        // Regex to extract src attributes from img tags
        Regex imgSrcRegex = new Regex("<img[^>]+src\\s*=\\s*\"([^\"]+)\"", RegexOptions.IgnoreCase);
        MatchCollection matches = imgSrcRegex.Matches(htmlContent);

        if (matches.Count == 0)
        {
            Console.WriteLine("No <img> tags found in the HTML. No chart images to verify.");
            return;
        }

        // Determine the folder where Aspose.Cells stores the image files
        string htmlDirectory = Path.GetDirectoryName(htmlFilePath) ?? string.Empty;
        string htmlFileNameWithoutExt = Path.GetFileNameWithoutExtension(htmlFilePath);
        string imageFolder = Path.Combine(htmlDirectory, $"{htmlFileNameWithoutExt}_files");

        bool allPng = true;

        foreach (Match match in matches)
        {
            string src = match.Groups[1].Value;
            // Resolve relative path to absolute path
            string imagePath = Path.Combine(htmlDirectory, src.Replace('/', Path.DirectorySeparatorChar));

            // If the src points to the image folder, adjust the path
            if (!Path.IsPathRooted(imagePath) && Directory.Exists(imageFolder))
            {
                imagePath = Path.Combine(imageFolder, Path.GetFileName(src));
            }

            string extension = Path.GetExtension(imagePath).ToLowerInvariant();

            Console.WriteLine($"Found image: '{src}' -> Extension: '{extension}'");

            if (extension != ".png")
            {
                allPng = false;
                Console.WriteLine($"[ERROR] Image '{src}' is not a PNG file.");
            }
            else if (!File.Exists(imagePath))
            {
                allPng = false;
                Console.WriteLine($"[ERROR] Expected PNG file '{imagePath}' does not exist on disk.");
            }
        }

        if (allPng)
        {
            Console.WriteLine("Verification succeeded: All chart images referenced in the HTML are PNG files.");
        }
        else
        {
            Console.WriteLine("Verification failed: Some chart images are not PNG files or missing.");
        }
    }
}
