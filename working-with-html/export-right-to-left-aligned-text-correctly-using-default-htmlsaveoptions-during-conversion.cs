// Title: Save Arabic right-to-left text to HTML using Aspose.Cells default HtmlSaveOptions in C#
// AI Prompts: Write C# code that inserts Arabic RTL text into a worksheet cell, sets the cell's horizontal alignment to Right, and saves the workbook as HTML using Aspose.Cells default HtmlSaveOptions. | Adapt the example to export Hebrew RTL text while keeping the correct text direction in the generated HTML file with Aspose.Cells. | Create a reusable C# method that takes any RTL string and a target cell address, applies right alignment, and exports the workbook to HTML using Aspose.Cells default settings.
// Common Searches: Aspose.Cells C# export Arabic cell to HTML preserving right-to-left alignment | How to keep RTL text direction when saving Excel as HTML with Aspose.Cells | C# Aspose.Cells default HtmlSaveOptions RTL support for Hebrew | Saving right-aligned Arabic text to HTML using Aspose.Cells .NET library
// Tags: Aspose.Cells export RTL text to HTML | C# set cell alignment right Aspose.Cells | HTML save options RTL formatting Aspose.Cells | preserve right-to-left direction in HTML output | save workbook as HTML with Arabic text

using System;
using System.IO;
using Aspose.Cells;

// // This program creates a new workbook, writes Arabic right‑to‑left text into cell A1, sets the cell's horizontal alignment to right, and saves the workbook as an HTML file using Aspose.Cells default HtmlSaveOptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert right‑to‑left (RTL) text into cell A1 (example Arabic text)
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("مرحبا بالعالم"); // "Hello World" in Arabic

            // Configure the cell style for RTL rendering
            Style rtlStyle = cell.GetStyle();
            // Enable right alignment (sufficient for most RTL display scenarios)
            rtlStyle.HorizontalAlignment = TextAlignmentType.Right;
            cell.SetStyle(rtlStyle);

            // Define output file path
            string outputPath = "output.html";

            // Save the workbook to HTML
            workbook.Save(outputPath, SaveFormat.Html);

            Console.WriteLine($"Workbook successfully saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
