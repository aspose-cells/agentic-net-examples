using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsImageMapDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string excelPath = "input.xlsx";

            // Output folder for images and HTML
            string outputFolder = "output";
            Directory.CreateDirectory(outputFolder);

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Configure image rendering options (TIFF, 96 DPI)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Tiff,
                OnePagePerSheet = true,
                HorizontalResolution = 96,
                VerticalResolution = 96
            };

            // StringBuilder to compose the final HTML
            StringBuilder html = new StringBuilder();
            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html><head><meta charset=\"UTF-8\"><title>Workbook Image Map</title></head><body>");

            // Iterate through each worksheet
            for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
            {
                Worksheet sheet = workbook.Worksheets[wsIndex];

                // Create a SheetRender for the current worksheet
                SheetRender sheetRender = new SheetRender(sheet, imgOptions);

                // Process each page of the worksheet
                for (int pageIndex = 0; pageIndex < sheetRender.PageCount; pageIndex++)
                {
                    // Determine image file name
                    string imageFileName = $"sheet{wsIndex}_page{pageIndex}.tiff";
                    string imagePath = Path.Combine(outputFolder, imageFileName);

                    // Render the page to a TIFF file
                    sheetRender.ToImage(pageIndex, imagePath); // uses ImageOrPrintOptions.ImageType = Tiff

                    // Get page size in inches and convert to pixels using DPI
                    float[] pageSizeInch = sheetRender.GetPageSizeInch(pageIndex);
                    int widthPx = (int)(pageSizeInch[0] * imgOptions.HorizontalResolution);
                    int heightPx = (int)(pageSizeInch[1] * imgOptions.VerticalResolution);

                    // Define a unique map name
                    string mapName = $"map_ws{wsIndex}_p{pageIndex}";

                    // Example hyperlink – in real scenarios this could be built from cell data
                    string hyperlink = $"https://example.com/resource?sheet={wsIndex}&page={pageIndex}";

                    // Add the image with usemap attribute
                    html.AppendLine($"<img src=\"{imageFileName}\" usemap=\"#{mapName}\" width=\"{widthPx}\" height=\"{heightPx}\" alt=\"Sheet {wsIndex} Page {pageIndex}\"/>");

                    // Create the image map covering the whole image
                    html.AppendLine($"<map name=\"{mapName}\">");
                    html.AppendLine($"  <area shape=\"rect\" coords=\"0,0,{widthPx},{heightPx}\" href=\"{hyperlink}\" target=\"_blank\" />");
                    html.AppendLine("</map>");
                    html.AppendLine("<br/>");
                }

                sheetRender.Dispose();
            }

            html.AppendLine("</body></html>");

            // Save the HTML file
            string htmlPath = Path.Combine(outputFolder, "workbook_image_map.html");
            File.WriteAllText(htmlPath, html.ToString());

            Console.WriteLine("Image map HTML generated at: " + htmlPath);
        }
    }
}