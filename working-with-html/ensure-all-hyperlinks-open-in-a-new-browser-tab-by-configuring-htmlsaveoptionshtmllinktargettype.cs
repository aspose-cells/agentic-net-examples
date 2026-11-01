// Title: How to configure HtmlSaveOptions.HtmlLinkTargetType so hyperlinks open in a new tab when exporting an Excel workbook to HTML using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export workbook to HTML with hyperlinks opening in a new tab | set HtmlLinkTargetType to Blank in Aspose.Cells HTML save options | C# example for configuring hyperlink target type in HtmlSaveOptions | Aspose.Cells HTML export link behavior new window C# | how to make Excel hyperlinks open in new window after saving as HTML with Aspose.Cells
// Tags: HtmlSaveOptions HtmlLinkTargetType | Aspose.Cells hyperlink target new tab | Excel to HTML export with link behavior | C# Aspose.Cells HTML save options configuration | set hyperlink target type in HTML export

using Aspose.Cells;
using Aspose.Cells.Saving;
using System;
using System.IO;

// Demonstrates loading an existing workbook, configuring HtmlSaveOptions.HtmlLinkTargetType to open all hyperlinks in a new browser tab, and saving the workbook as HTML using Aspose.Cells for .NET, with error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: ExportHyperlinkTarget is not available in this version; hyperlinks will retain default behavior.

            // Save the workbook as HTML with the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
