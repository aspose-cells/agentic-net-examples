// Title: Convert an Excel workbook to HTML in C# with Aspose.Cells while keeping original numeric formatting
// Common Searches: C# Aspose.Cells export Excel to HTML preserving number formats | How to retain Excel cell numeric formatting when saving as HTML using Aspose.Cells | Aspose.Cells HtmlSaveOptions keep original number display in HTML output | Save workbook as HTML with exact numeric formatting in C# | Preserve Excel numeric formatting in generated HTML file with Aspose.Cells library
// Tags: Aspose.Cells HtmlSaveOptions numeric formatting | C# Excel to HTML conversion preserving number format | export Excel workbook as HTML with original formatting | Aspose.Cells retain cell display in HTML output | HTML export of Excel with numeric format retention

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example loads an Excel file (default "input.xlsx" or a path supplied via command line), creates a Workbook object, configures HtmlSaveOptions, and saves the workbook as "output.html". The default Aspose.Cells HTML export preserves the original numeric formatting of cells, while handling missing files and conversion errors.
class Program
{
    static void Main(string[] args)
    {
        // Determine the input Excel file path.
        // If a path is supplied as a command‑line argument it is used,
        // otherwise the program expects a file named "input.xlsx" in the same folder as the executable.
        string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
        string outputPath = "output.html";

        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The input file \"{inputPath}\" does not exist.");
            return;
        }

        try
        {
            // Load the workbook from the Excel file.
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as HTML using the configured options.
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Successfully converted \"{Path.GetFileName(inputPath)}\" to HTML.");
            Console.WriteLine($"HTML file saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while converting the workbook to HTML:");
            Console.WriteLine(ex.Message);
        }
    }
}
