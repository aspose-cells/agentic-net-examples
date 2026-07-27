using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Sample");
        worksheet.Cells["B1"].PutValue("Data");

        // Configure image rendering options for PNG output
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png
        };

        // Render the first page of the worksheet to a PNG file
        SheetRender sheetRender = new SheetRender(worksheet, imgOptions);
        string imageFile = "worksheet.png";
        sheetRender.ToImage(0, imageFile); // page index 0

        // Create markdown content that references the generated PNG image
        string markdownContent = $"![Worksheet Image]({imageFile})";

        // Save the markdown document
        string markdownFile = "worksheet.md";
        File.WriteAllText(markdownFile, markdownContent);

        Console.WriteLine($"Markdown file '{markdownFile}' created with image reference '{imageFile}'.");
    }
}