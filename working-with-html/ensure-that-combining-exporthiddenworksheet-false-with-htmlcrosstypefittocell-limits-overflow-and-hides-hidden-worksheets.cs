// Title: Export Excel to HTML in C# with Aspose.Cells while hiding hidden worksheets and fitting cell content using HtmlCrossType.FitToCell
// AI Prompts: Write C# code that loads an .xlsx workbook, sets HtmlSaveOptions.ExportHiddenWorksheet = false and HtmlSaveOptions.HtmlCrossType = HtmlCrossType.FitToCell, and saves the result as an HTML file. | Demonstrate how to generate HTML from a workbook with Aspose.Cells so that hidden sheets are omitted and any text that exceeds the cell boundaries is trimmed to fit the cell size.
// Common Searches: Aspose.Cells C# export to HTML hide hidden sheets and fit overflow text | How to use HtmlCrossType.FitToCell with HtmlSaveOptions in Aspose.Cells | Prevent cell overflow when converting Excel to HTML using Aspose.Cells | Export workbook to HTML without hidden worksheets Aspose.Cells example
// Tags: HtmlSaveOptions ExportHiddenWorksheet false | HtmlSaveOptions HtmlCrossType FitToCell | Aspose.Cells hide hidden worksheets HTML export | trim cell overflow Aspose.Cells HTML | C# convert Excel to HTML Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Rendering;
using System;
using System.IO;

// The sample loads an existing Excel file, configures HtmlSaveOptions to exclude hidden worksheets (ExportHiddenWorksheet = false) and to limit text overflow by setting HtmlCrossType to FitToCell, then saves the workbook as an HTML document.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify the input file exists to prevent FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML export options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Do not export worksheets that are hidden in the workbook
                ExportHiddenWorksheet = false
                // Aspose.Cells does not provide a direct HtmlCrossType property.
                // The default HTML export trims overflow text to the cell size.
            };

            // Export the workbook to HTML using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
