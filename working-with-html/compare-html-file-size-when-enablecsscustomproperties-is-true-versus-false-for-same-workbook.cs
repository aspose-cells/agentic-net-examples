// Title: C# Sample: Compare Aspose.Cells HTML Export Size with EnableCssCustomProperties True vs False
// Description: This example creates a workbook with styled text and two identical images, saves it to HTML twice—once with HtmlSaveOptions.EnableCssCustomProperties enabled and once disabled—then prints the byte size of each file to show which setting produces a smaller output.
// Keywords: Aspose.Cells HTML export | EnableCssCustomProperties | .NET C# example | HTML file size comparison | CSS custom properties performance | lightweight HTML report | Aspose.Cells sample code | GitHub Aspose.Cells example | benchmark HTML output
// Common Searches: Aspose.Cells EnableCssCustomProperties file size | HTML export size difference Aspose.Cells | C# compare HTML output with and without CSS custom properties | reduce HTML size using Aspose.Cells | measure Aspose.Cells HTML export size
// Developer Intent: Find out whether turning on CSS custom properties shrinks the HTML generated from a workbook.
// Use Cases: Create compact HTML reports for web publishing by toggling EnableCssCustomProperties. | Benchmark the impact of CSS custom properties on HTML size for workbooks containing styled cells and images. | Select optimal HtmlSaveOptions settings for performance‑critical web applications.
// AI Prompts: Generate C# code that logs the byte difference between two HTML files saved with EnableCssCustomProperties set to true and false using Aspose.Cells. | Explain how enabling CSS custom properties changes the generated markup and influences file size in Aspose.Cells HTML export. | Recommend additional HtmlSaveOptions that can further reduce HTML output size while keeping formatting intact.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsHtmlSizeComparison
{
    // This example creates a workbook with styled text and two identical images, saves it to HTML twice—once with HtmlSaveOptions.EnableCssCustomProperties enabled and once disabled—then prints the byte size of each file to show which setting produces a smaller output.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add some sample data
                worksheet.Cells["A1"].PutValue("Sample Text");
                Style style = worksheet.Cells["A1"].GetStyle();
                style.Font.IsBold = true;
                style.Font.Color = Color.Red;
                worksheet.Cells["A1"].SetStyle(style);

                // Path to the image file (ensure it exists)
                string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logo.png");
                if (File.Exists(imagePath))
                {
                    // Add the same image to two different cells to demonstrate CSS custom properties effect
                    int imgIndex1 = worksheet.Pictures.Add(1, 1, imagePath);
                    Picture pic1 = worksheet.Pictures[imgIndex1];
                    pic1.Width = 100;
                    pic1.Height = 100;

                    int imgIndex2 = worksheet.Pictures.Add(5, 3, imagePath);
                    Picture pic2 = worksheet.Pictures[imgIndex2];
                    pic2.Width = 100;
                    pic2.Height = 100;
                }
                else
                {
                    Console.WriteLine($"Image file not found: {imagePath}. Skipping image insertion.");
                }

                // Prepare HtmlSaveOptions instance
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

                // Save with EnableCssCustomProperties = true
                htmlOptions.EnableCssCustomProperties = true;
                string fileWithCss = "HtmlWithCssCustomProperties.html";
                workbook.Save(fileWithCss, htmlOptions);

                // Save with EnableCssCustomProperties = false
                htmlOptions.EnableCssCustomProperties = false;
                string fileWithoutCss = "HtmlWithoutCssCustomProperties.html";
                workbook.Save(fileWithoutCss, htmlOptions);

                // Compare file sizes
                long sizeWithCss = new FileInfo(fileWithCss).Length;
                long sizeWithoutCss = new FileInfo(fileWithoutCss).Length;

                Console.WriteLine($"File size with CSS custom properties: {sizeWithCss} bytes");
                Console.WriteLine($"File size without CSS custom properties: {sizeWithoutCss} bytes");

                if (sizeWithCss < sizeWithoutCss)
                {
                    Console.WriteLine("Enabling CSS custom properties reduces the HTML file size.");
                }
                else if (sizeWithCss > sizeWithoutCss)
                {
                    Console.WriteLine("Disabling CSS custom properties results in a smaller HTML file.");
                }
                else
                {
                    Console.WriteLine("Both files have the same size.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
