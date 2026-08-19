// Title: Embed a Base64 PNG in a Merged Excel Cell using Aspose.Cells HtmlString (C#)
// Description: Creates a workbook, merges B2:D2, reads a PNG, converts it to a Base64 data‑URI, assigns the <img> tag to the merged cell via HtmlString, and saves the file so the image appears inside the merged range when opened.
// Keywords: Aspose.Cells HtmlString | C# embed image in Excel cell | Base64 data URI Excel | merged cells picture Aspose | insert PNG into merged cell | Excel image rendering C# | Aspose.Cells image merge
// Common Searches: Aspose.Cells embed image in merged cell C# | HtmlString property base64 image Excel | display picture after merging cells Aspose | C# add PNG to merged Excel range | how to use data‑uri image in Aspose.Cells
// Developer Intent: Insert an image that spans a merged cell range by setting the cell’s HtmlString to a Base64‑encoded <img> tag.
// Use Cases: Add a company logo across a merged header row in automated reports. | Show product thumbnails in merged catalog cells generated programmatically. | Place dynamically created chart snapshots in merged dashboard sections for scheduled exports.
// AI Prompts: Modify the example to embed a JPEG image and specify width/height attributes. | Show how to center the image within the merged cell using CSS styles in the HtmlString. | Provide error‑handling code that creates a placeholder cell when the image file is missing.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, merges B2:D2, reads a PNG, converts it to a Base64 data‑URI, assigns the <img> tag to the merged cell via HtmlString, and saves the file so the image appears inside the merged range when opened.
    public class HtmlImageInMergedCellDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Merge cells B2:D2 (row index 1, columns 1 to 3)
                sheet.Cells.Merge(1, 1, 1, 3);

                // Load an image file (replace with your image path)
                string imagePath = "sample.png";
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {Path.GetFullPath(imagePath)}");
                    return;
                }

                byte[] imageBytes = File.ReadAllBytes(imagePath);

                // Convert image bytes to a Base64 string
                string base64 = Convert.ToBase64String(imageBytes);

                // Build the HTML img tag using a data URI
                string htmlImgTag = $"<img src=\"data:image/png;base64,{base64}\" />";

                // Set the HTML string of the merged cell (top‑left cell of the range)
                Cell mergedCell = sheet.Cells[1, 1]; // B2
                mergedCell.HtmlString = htmlImgTag;

                // Save the workbook; the image will be rendered inside the merged cell when opened
                string outputPath = "HtmlImageInMergedCell.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            HtmlImageInMergedCellDemo.Run();
        }
    }
}
