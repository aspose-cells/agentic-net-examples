// Title: Aspose.Cells .NET: Compare HTML size with EnableCssCustomProperties true vs false
// Description: Creates a workbook with styled text and optional images, saves it twice as HTML—once with HtmlSaveOptions.EnableCssCustomProperties enabled and once disabled—then reports the byte size of each file and the difference.
// Keywords: Aspose.Cells HTML size | EnableCssCustomProperties | HtmlSaveOptions performance | C# HTML export size comparison | Aspose.Cells CSS custom properties impact
// Common Searches: Aspose.Cells EnableCssCustomProperties file size | HTML export size difference Aspose.Cells .NET | Does CSS custom properties increase HTML output size | Measure Aspose.Cells HTML size with and without custom properties
// Developer Intent: Find out how toggling EnableCssCustomProperties influences the generated HTML file size.
// Use Cases: Assess storage or bandwidth savings by disabling CSS custom properties for web‑served spreadsheets. | Choose optimal HtmlSaveOptions for performance‑critical applications that embed images and styled cells. | Create automated reports that log HTML export sizes for different configuration settings.
// AI Prompts: Generate a C# program that saves a workbook to HTML with EnableCssCustomProperties set to true and false, then prints both file sizes. | Explain why enabling CSS custom properties can enlarge the HTML output in Aspose.Cells and suggest techniques to minimize the size. | Write a PowerShell script that executes the compiled C# app, captures the size metrics, and appends them to a CSV log.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Required for Picture class

namespace AsposeCellsHtmlSizeComparison
{
    // Creates a workbook with styled text and optional images, saves it twice as HTML—once with HtmlSaveOptions.EnableCssCustomProperties enabled and once disabled—then reports the byte size of each file and the difference.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add some sample data with formatting
                var cellA1 = sheet.Cells["A1"];
                cellA1.PutValue("Sample Text");
                var style = cellA1.GetStyle();
                style.Font.IsBold = true;
                style.Font.Color = System.Drawing.Color.Red;
                cellA1.SetStyle(style);

                // Path to the image file
                string imagePath = "logo.png";

                // Add image if it exists; otherwise skip image insertion
                if (File.Exists(imagePath))
                {
                    // Add the image the first time
                    int imgIndex1 = sheet.Pictures.Add(1, 1, imagePath);
                    Picture pic1 = sheet.Pictures[imgIndex1];
                    pic1.Width = 100;
                    pic1.Height = 100;

                    // Add the same image again in another cell
                    int imgIndex2 = sheet.Pictures.Add(5, 3, imagePath);
                    Picture pic2 = sheet.Pictures[imgIndex2];
                    pic2.Width = 100;
                    pic2.Height = 100;
                }
                else
                {
                    Console.WriteLine($"Image file \"{imagePath}\" not found. Skipping image insertion.");
                }

                // Prepare HTML save options for the first file (EnableCssCustomProperties = true)
                HtmlSaveOptions optionsWithCssCustom = new HtmlSaveOptions
                {
                    EnableCssCustomProperties = true
                };
                string fileWithCssCustom = "Html_WithCssCustomProperties.html";
                workbook.Save(fileWithCssCustom, optionsWithCssCustom);

                // Prepare HTML save options for the second file (EnableCssCustomProperties = false)
                HtmlSaveOptions optionsWithoutCssCustom = new HtmlSaveOptions
                {
                    EnableCssCustomProperties = false
                };
                string fileWithoutCssCustom = "Html_WithoutCssCustomProperties.html";
                workbook.Save(fileWithoutCssCustom, optionsWithoutCssCustom);

                // Get file sizes
                long sizeWith = new FileInfo(fileWithCssCustom).Length;
                long sizeWithout = new FileInfo(fileWithoutCssCustom).Length;

                // Output the comparison results
                Console.WriteLine($"File size with EnableCssCustomProperties=true : {sizeWith} bytes");
                Console.WriteLine($"File size with EnableCssCustomProperties=false: {sizeWithout} bytes");
                Console.WriteLine($"Size difference: {Math.Abs(sizeWith - sizeWithout)} bytes");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
