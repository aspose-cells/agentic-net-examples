// Title: C# – Embed a Base64 PNG in a Merged Cell with Aspose.Cells HtmlString
// Description: Demonstrates how to merge cells, assign a Base64‑encoded <img> tag to the top‑left cell via the HtmlString property, and save the workbook as HTML (images embedded) and XLSX (HTML markup retained) using Aspose.Cells for .NET.
// Keywords: Aspose.Cells HtmlString | embed Base64 image Excel | merged cell image C# | export workbook to HTML with images | preserve HTML tags in XLSX | Aspose.Cells .NET image handling
// Common Searches: Aspose.Cells set HtmlString for merged cell | C# embed Base64 PNG in Excel cell | how to display image in merged Excel cells | export Aspose.Cells workbook to HTML with embedded images | keep <img> tag when saving XLSX with Aspose.Cells
// Developer Intent: Insert an <img> tag into a merged cell so the picture appears after merging and is retained in HTML and XLSX outputs.
// Use Cases: Add a logo or banner to a merged header row without external image files. | Generate HTML reports from spreadsheets that contain inline Base64 images. | Create XLSX files that preserve HTML image markup for downstream applications supporting HtmlString.
// AI Prompts: Show C# code that merges cells, sets a Base64 PNG via HtmlString, and saves to HTML with images embedded. | Provide an Aspose.Cells example that keeps an <img> tag in a merged cell when exporting to XLSX. | Explain HtmlSaveOptions settings needed to retain Base64 images while exporting a workbook containing HtmlString content.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to merge cells, assign a Base64‑encoded <img> tag to the top‑left cell via the HtmlString property, and save the workbook as HTML (images embedded) and XLSX (HTML markup retained) using Aspose.Cells for .NET.
    public class HtmlImageInMergedCellDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Merge a range of cells (A1:D2) – 2 rows, 4 columns
            worksheet.Cells.Merge(0, 0, 2, 4);

            // Small transparent PNG image encoded as Base64 (1x1 pixel)
            const string pngBase64 =
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO2b4ZcAAAAASUVORK5CYII=";

            // Build the HTML img tag using the Base64 data
            string imgTag = $"<img src=\"data:image/png;base64,{pngBase64}\" alt=\"Embedded\"/>";

            // Set the HTML string of the merged cell (top‑left cell of the merged range)
            Cell mergedCell = worksheet.Cells["A1"];
            mergedCell.HtmlString = imgTag;

            // Export the workbook to HTML to verify rendering
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = true // keep images embedded in HTML
            };
            string htmlPath = "MergedCellWithImage.html";
            workbook.Save(htmlPath, htmlOptions);
            Console.WriteLine($"HTML file saved to: {Path.GetFullPath(htmlPath)}");

            // Save the workbook in XLSX format (the HTML string is preserved)
            string xlsxPath = "MergedCellWithImage.xlsx";
            workbook.Save(xlsxPath);
            Console.WriteLine($"XLSX file saved to: {Path.GetFullPath(xlsxPath)}");
        }
    }
}
