// Title: Convert an Excel workbook to HTML with default settings and include cell comments using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that checks for an .xlsx file, loads it with Aspose.Cells, and saves it as HTML while preserving cell comments. | Show how to use HtmlSaveOptions with SaveFormat.Html in Aspose.Cells to export a workbook to HTML with default options and comment support. | Create a robust console application in C# that converts an Excel workbook to HTML, includes error handling, and ensures comments are exported.
// Common Searches: Aspose.Cells C# export Excel to HTML with cell comments included | How to save a workbook as HTML preserving comments using Aspose.Cells .NET | C# convert .xlsx to .html default HtmlSaveOptions Aspose.Cells example | Example code for Excel to HTML conversion with comments in Aspose.Cells for .NET
// Tags: Aspose.Cells HtmlSaveOptions default export | C# Excel to HTML conversion with comments | Export cell comments Aspose.Cells HTML | Workbook.Save HTML Aspose.Cells example | File existence check error handling Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program checks that the input Excel file exists, loads it into an Aspose.Cells Workbook, creates HtmlSaveOptions with default settings (which include cell comments), and saves the workbook as an HTML file while handling any exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (cell comments are exported by default in recent versions)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
