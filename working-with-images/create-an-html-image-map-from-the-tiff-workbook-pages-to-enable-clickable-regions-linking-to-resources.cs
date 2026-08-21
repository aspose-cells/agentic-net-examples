// Title: C# – Render Excel Worksheets to PNG and Generate an HTML Image Map with Aspose.Cells
// Description: Sample code that creates a workbook, renders each worksheet to a PNG image using Aspose.Cells WorkbookRender (96 DPI), computes pixel dimensions from the page size, and produces an HTML file containing <img> tags linked to full‑size <area> elements. Each image map points to a placeholder resource URL, demonstrating how to add clickable regions to worksheet previews.
// Keywords: Aspose.Cells C# render worksheet to image | WorkbookRender PNG output | HTML image map from Excel sheet | clickable area Aspose.Cells | Excel to PNG with map C# | Aspose.Cells ImageOrPrintOptions | generate HTML map for workbook pages
// Common Searches: Aspose.Cells create image map from Excel sheets C# | render each worksheet as PNG and add clickable area | C# code to generate HTML image map for Excel workbook | how to use WorkbookRender for image maps | convert Excel pages to images with Aspose.Cells
// Developer Intent: Produce an HTML page that displays each worksheet as an image and attaches a full‑size clickable region linking to a specific URL.
// Use Cases: Show worksheet previews on a web portal where clicking a page opens a detailed report. | Embed Excel sheet images in documentation with links to related help topics. | Create a knowledge‑base navigation where each sheet image maps to a tutorial or FAQ.
// AI Prompts: Write C# code using Aspose.Cells to render every worksheet to PNG and build an HTML image map that links each image to a custom URL. | Explain how to convert the page size returned by WorkbookRender (in inches) to pixel coordinates for accurate image‑map areas. | Modify the example to output TIFF files instead of PNG while keeping the generated HTML image map functional.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsImageMapDemo
{
    // Sample code that creates a workbook, renders each worksheet to a PNG image using Aspose.Cells WorkbookRender (96 DPI), computes pixel dimensions from the page size, and produces an HTML file containing <img> tags linked to full‑size <area> elements. Each image map points to a placeholder resource URL, demonstrating how to add clickable regions to worksheet previews.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a sample workbook with two worksheets
                Workbook workbook = new Workbook();

                // First worksheet (already exists at index 0)
                Worksheet ws1 = workbook.Worksheets[0];
                ws1.Name = "FirstSheet";
                ws1.Cells["A1"].PutValue("Page 1 - Clickable Area");
                ws1.Cells["A2"].PutValue("Data...");

                // Add second worksheet
                int ws2Index = workbook.Worksheets.Add();
                Worksheet ws2 = workbook.Worksheets[ws2Index];
                ws2.Name = "SecondSheet";
                ws2.Cells["A1"].PutValue("Page 2 - Clickable Area");
                ws2.Cells["A2"].PutValue("More Data...");

                // Configure image rendering options (default format is PNG)
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    HorizontalResolution = 96, // DPI for width calculation
                    VerticalResolution = 96,   // DPI for height calculation
                    OnePagePerSheet = true     // One image per sheet page
                };

                // Initialize workbook renderer
                WorkbookRender renderer = new WorkbookRender(workbook, imgOptions);

                // Prepare HTML builder
                StringBuilder htmlBuilder = new StringBuilder();
                htmlBuilder.AppendLine("<!DOCTYPE html>");
                htmlBuilder.AppendLine("<html>");
                htmlBuilder.AppendLine("<head><meta charset=\"UTF-8\"><title>Workbook Image Map</title></head>");
                htmlBuilder.AppendLine("<body>");

                // Iterate through each rendered page
                for (int i = 0; i < renderer.PageCount; i++)
                {
                    try
                    {
                        // Define file name for the image of the current page
                        string imageFileName = $"page_{i}.png";

                        // Render the page to an image file
                        renderer.ToImage(i, imageFileName);

                        // Obtain page size in inches and convert to pixels using the DPI set above
                        float[] pageSizeInch = renderer.GetPageSizeInch(i);
                        int widthPx = (int)(pageSizeInch[0] * imgOptions.HorizontalResolution);
                        int heightPx = (int)(pageSizeInch[1] * imgOptions.VerticalResolution);

                        // Build <img> tag referencing the rendered image and associate it with a map
                        string mapName = $"map{i}";
                        htmlBuilder.AppendLine($"<img src=\"{imageFileName}\" usemap=\"#{mapName}\" width=\"{widthPx}\" height=\"{heightPx}\" alt=\"Page {i}\"/>");

                        // Create an image map covering the whole image; replace href with actual resource URL as needed
                        string resourceUrl = $"resource_page_{i}.html"; // Placeholder URL
                        htmlBuilder.AppendLine($"<map name=\"{mapName}\">");
                        htmlBuilder.AppendLine($"  <area shape=\"rect\" coords=\"0,0,{widthPx},{heightPx}\" href=\"{resourceUrl}\" alt=\"Link to resource {i}\"/>");
                        htmlBuilder.AppendLine("</map>");
                        htmlBuilder.AppendLine("<br/>"); // Separate images visually
                    }
                    catch (Exception exPage)
                    {
                        Console.Error.WriteLine($"Error processing page {i}: {exPage.Message}");
                    }
                }

                htmlBuilder.AppendLine("</body>");
                htmlBuilder.AppendLine("</html>");

                // Save the generated HTML to a file
                string htmlOutputPath = "WorkbookImageMap.html";
                File.WriteAllText(htmlOutputPath, htmlBuilder.ToString(), Encoding.UTF8);

                Console.WriteLine("Image pages and HTML image map have been generated:");
                Console.WriteLine($"- Images: page_0.png ... page_{renderer.PageCount - 1}.png");
                Console.WriteLine($"- HTML: {htmlOutputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
