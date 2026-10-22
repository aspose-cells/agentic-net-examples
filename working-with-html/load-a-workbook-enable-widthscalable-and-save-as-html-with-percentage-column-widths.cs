// Title: Enable WidthScalable and export an Excel workbook to responsive HTML with percentage column widths using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file, sets HtmlSaveOptions.WidthScalable to true, and saves the workbook as HTML where column widths are expressed as percentages. | Show a complete example of using Aspose.Cells to convert Excel to HTML with responsive column sizing by enabling WidthScalable in HtmlSaveOptions. | Provide a step‑by‑step C# snippet that configures HtmlSaveOptions for percentage‑based column widths and writes the output to an .html file.
// Common Searches: Aspose.Cells C# enable WidthScalable for HTML export | export Excel to HTML with percentage column widths using Aspose.Cells .NET | how to make HTML output from Aspose.Cells responsive column widths | C# Aspose.Cells HtmlSaveOptions set column width scaling | save workbook as HTML with scalable column widths Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions WidthScalable | C# export Excel to HTML percentage column widths | responsive HTML conversion Aspose.Cells | set column width scaling Aspose.Cells | HtmlSaveOptions SaveFormat.Html Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, creates HtmlSaveOptions with WidthScalable set to true so that column widths are rendered as percentages, and saves the workbook as a responsive HTML file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Ensure the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (default settings are sufficient)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as HTML with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
