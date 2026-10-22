// Title: Render the first worksheet of an Excel workbook to a PNG file and embed it in a Markdown document using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, renders the first worksheet to a PNG image, and creates a Markdown file containing a ![alt](path) image tag that references the PNG. | Modify the program to loop through all worksheets, save each as a separate PNG, and generate a single Markdown file that lists every image with a caption showing the worksheet name.
// Common Searches: how to use Aspose.Cells in C# to convert an Excel sheet to PNG | C# code to embed a generated PNG from Excel into a markdown file | export multiple Excel worksheets as images and create markdown documentation with Aspose.Cells | Aspose.Cells SheetRender one page per sheet example C# | generate markdown image links from Excel data using .NET
// Tags: Aspose.Cells SheetRender PNG export | C# render worksheet to image | markdown embed image path C# | one-page-per-sheet option Aspose.Cells | generate markdown from Excel images

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads 'input.xlsx' with Aspose.Cells, uses SheetRender to export the first worksheet as 'worksheet.png', and writes 'document.md' containing a Markdown image tag that points to the PNG file.
class Program
{
    static void Main()
    {
        // Define file paths
        string excelPath = "input.xlsx";
        string pngPath = "worksheet.png";
        string markdownPath = "document.md";

        // Verify that the source Excel file exists
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Error: Excel file not found at '{excelPath}'.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure image export options
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Default image format is PNG; no need to set explicitly
                OnePagePerSheet = true // One page per sheet
            };

            // Render the worksheet to a PNG image
            SheetRender sheetRender = new SheetRender(worksheet, imgOptions);
            sheetRender.ToImage(0, pngPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during Excel processing or image export: {ex.Message}");
            return;
        }

        try
        {
            // Build Markdown content that embeds the PNG image
            string markdownContent = $"![Worksheet Image]({pngPath})";

            // Write the Markdown content to the file
            File.WriteAllText(markdownPath, markdownContent);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error writing Markdown file: {ex.Message}");
            return;
        }

        Console.WriteLine("Export completed successfully.");
    }
}
