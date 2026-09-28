// Title: Render an Excel worksheet to a PNG file and embed it in a Markdown document using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, saves the first worksheet as a PNG image with Aspose.Cells, and writes a Markdown file containing an ![alt](path) link to the image. | Write C# that iterates over all worksheets in a workbook, exports each to a separate PNG file using Aspose.Cells, and builds a single Markdown document that lists the images sequentially. | Provide a C# example that adds robust try‑catch logging, ensures output directories exist, and produces both PNG images of worksheets and a Markdown file with embedded image links via Aspose.Cells.
// Common Searches: Aspose.Cells C# export worksheet as PNG and create markdown link | how to save Excel sheet as image and embed in markdown using .NET | C# code to generate markdown file with worksheet screenshot from Aspose.Cells | render multiple Excel worksheets to PNG and combine into one markdown document in C#
// Tags: Aspose.Cells render worksheet to PNG | C# export Excel sheet as image | generate markdown with embedded image C# | SheetRender PNG export Aspose.Cells | markdown image link from Excel worksheet

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample creates a new workbook, renders the first worksheet to a PNG file using Aspose.Cells, and writes a Markdown file that embeds the PNG image with the standard ![alt](path) syntax.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sample";

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Define the PNG file path where the worksheet image will be saved
            string pngPath = "WorksheetImage.png";

            // Ensure the directory for the PNG exists
            string pngDir = Path.GetDirectoryName(pngPath);
            if (!string.IsNullOrEmpty(pngDir) && !Directory.Exists(pngDir))
            {
                Directory.CreateDirectory(pngDir);
            }

            // Set image export options (one page per sheet). PNG is the default format.
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Render the worksheet to an image and save it as PNG
            SheetRender sheetRender = new SheetRender(sheet, imgOptions);
            sheetRender.ToImage(0, pngPath); // 0 = first page of the sheet

            // Build the markdown string that embeds the PNG image
            string markdownContent = $"![Worksheet Image]({pngPath})";

            // Define the markdown file path
            string mdPath = "WorksheetImage.md";

            // Ensure the directory for the markdown file exists
            string mdDir = Path.GetDirectoryName(mdPath);
            if (!string.IsNullOrEmpty(mdDir) && !Directory.Exists(mdDir))
            {
                Directory.CreateDirectory(mdDir);
            }

            // Save the markdown content to a .md file
            File.WriteAllText(mdPath, markdownContent);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
