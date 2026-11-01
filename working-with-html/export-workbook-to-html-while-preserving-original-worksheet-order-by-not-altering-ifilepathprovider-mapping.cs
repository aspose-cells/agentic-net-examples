// Title: Export an Excel workbook to HTML while keeping the original worksheet order using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export workbook to HTML preserving sheet order | How to save Excel as HTML without reordering worksheets in Aspose.Cells | Aspose.Cells HtmlSaveOptions keep original worksheet sequence | Save Excel file to HTML with same sheet arrangement using .NET | Prevent worksheet order change when converting .xlsx to HTML with Aspose
// Tags: Aspose.Cells HtmlSaveOptions preserve worksheet order | C# export Excel to HTML Aspose.Cells | save workbook as HTML original sheet sequence | Excel to HTML conversion maintaining sheet order | Aspose.Cells SaveFormat.Html worksheet ordering

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates loading an existing .xlsx file with Aspose.Cells, using HtmlSaveOptions (default behavior) to retain the original worksheet sequence, and saving the workbook as an HTML file while handling missing file and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // No additional settings required to preserve worksheet order; default behavior retains it.

            // Export the workbook to HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
