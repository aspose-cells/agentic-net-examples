// Title: Convert an Excel workbook with DataBar conditional formatting to HTML while preserving DataBar visuals using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file containing DataBar conditional formatting, configures HtmlSaveOptions to retain the formatting, and saves the workbook as an HTML file with Aspose.Cells. | Show how to add file‑existence validation and comprehensive exception handling to an Aspose.Cells Excel‑to‑HTML conversion that keeps DataBar visuals intact. | Demonstrate enabling DataBar export in Aspose.Cells and verify that the generated HTML correctly displays the DataBar bars.
// Common Searches: Aspose.Cells C# export Excel with DataBar conditional formatting to HTML | How to keep DataBar bars when converting .xlsx to HTML using Aspose.Cells | HtmlSaveOptions preserve conditional formatting DataBars Aspose.Cells .NET | C# example converting workbook with DataBars to HTML | Save Excel workbook as HTML with visual DataBar bars using Aspose.Cells
// Tags: Aspose.Cells HTML export with DataBar conditional formatting | C# HtmlSaveOptions preserve DataBar visuals | Excel to HTML conversion retaining conditional formatting | DataBar export in Aspose.Cells .NET | load workbook and save as HTML Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an existing .xlsx workbook that contains DataBar conditional formatting, checks that the file exists, creates a Workbook object, uses default HtmlSaveOptions (which preserve conditional formatting including DataBars), and saves the workbook as an HTML file while maintaining the visual appearance of the DataBars.
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

            // Load the workbook that contains conditional formatting with DataBars
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (default settings already preserve conditional formatting)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

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
