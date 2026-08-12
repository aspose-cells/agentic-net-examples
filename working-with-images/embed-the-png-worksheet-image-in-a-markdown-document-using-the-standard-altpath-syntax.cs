// Title: Render an Excel worksheet to PNG and embed it in Markdown using Aspose.Cells for .NET
// Description: A C# example that creates a workbook, adds sample data, saves it, renders the first worksheet as a PNG image with Aspose.Cells, and generates a Markdown file that embeds the image using the standard `![alt](path)` syntax.
// Keywords: Aspose.Cells PNG export | C# render worksheet to image | embed Excel image in Markdown | generate Markdown screenshot from Excel | Aspose.Cells markdown documentation | Excel to PNG conversion .NET | automated README image generation
// Common Searches: export Excel sheet as PNG with Aspose.Cells | C# code to embed worksheet image in README.md | how to create markdown image link from Excel screenshot | Aspose.Cells render worksheet to image and write markdown | automate markdown documentation with Excel screenshots
// Developer Intent: Produce a PNG snapshot of an Excel worksheet and automatically insert it into a Markdown file via the `![alt](path)` syntax.
// Use Cases: Generate up‑to‑date documentation that includes visual snapshots of Excel data. | Add worksheet images to GitHub README or other markdown‑based project pages. | Create CI/CD pipelines that output markdown reports with embedded Excel screenshots. | Build static‑site content where Excel tables are displayed as images for consistent styling.
// AI Prompts: Write C# code using Aspose.Cells to convert the first worksheet of a workbook to a PNG file and create a Markdown file that embeds the image with `![Worksheet Image](filename)`. | Explain how to configure ImageOrPrintOptions (resolution, scaling, margins) for high‑quality PNG output suitable for markdown documentation. | Provide a step‑by‑step guide to loop through all worksheets in an Excel file, render each to PNG, and generate a markdown file that includes all images with appropriate captions.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// A C# example that creates a workbook, adds sample data, saves it, renders the first worksheet as a PNG image with Aspose.Cells, and generates a Markdown file that embeds the image using the standard `![alt](path)` syntax.
class Program
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data
        worksheet.Cells["A1"].PutValue("Hello");
        worksheet.Cells["B1"].PutValue("World");

        // Optional: add a picture to the worksheet
        // int picIdx = worksheet.Pictures.Add(1, 1, "sample.png");

        // Save the workbook (save rule) – not strictly required for the markdown, but follows lifecycle
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);
        string workbookPath = Path.Combine(outputDir, "workbook.xlsx");
        workbook.Save(workbookPath);

        // Render the worksheet to a PNG image file
        string imagePath = Path.Combine(outputDir, "worksheet.png");
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = Aspose.Cells.Drawing.ImageType.Png,
            OnePagePerSheet = true
        };
        SheetRender sheetRender = new SheetRender(worksheet, imgOptions);
        // Render the first (and only) page to the PNG file
        sheetRender.ToImage(0, imagePath);

        // Create a Markdown file that embeds the PNG using the standard syntax
        string markdownPath = Path.Combine(outputDir, "worksheet.md");
        string markdownContent = $"![Worksheet Image]({Path.GetFileName(imagePath)})";
        File.WriteAllText(markdownPath, markdownContent);

        Console.WriteLine($"Workbook saved to: {workbookPath}");
        Console.WriteLine($"Worksheet image saved to: {imagePath}");
        Console.WriteLine($"Markdown file created at: {markdownPath}");
    }
}
