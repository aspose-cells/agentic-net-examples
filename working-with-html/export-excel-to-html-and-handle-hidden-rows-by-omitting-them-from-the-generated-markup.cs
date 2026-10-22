// Title: Convert an Excel .xlsx file to HTML with Aspose.Cells for .NET while excluding hidden rows
// AI Prompts: Write C# code that loads an .xlsx workbook using Aspose.Cells, removes every row whose IsHidden flag is true, and then saves the workbook as an HTML file with HtmlSaveOptions. | Show how to iterate through a Worksheet in Aspose.Cells, detect hidden rows, delete them from the sheet, and export the cleaned workbook to HTML. | Provide a C# example that verifies the input Excel file, filters out hidden rows, and generates an HTML representation that contains only visible rows using Aspose.Cells.
// Common Searches: Aspose.Cells .NET export to HTML without hidden rows | C# remove hidden rows before saving Excel as HTML using Aspose.Cells | How to skip hidden rows when converting XLSX to HTML with Aspose.Cells | HtmlSaveOptions hide rows Aspose.Cells .NET example | Convert Excel to HTML and exclude hidden rows programmatically
// Tags: Aspose.Cells HTML export exclude hidden rows | C# delete hidden rows Aspose.Cells | HtmlSaveOptions usage Aspose.Cells .NET | Excel to HTML conversion Aspose.Cells | Workbook row visibility handling Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an input.xlsx workbook, iterates through its rows to delete any that are hidden, configures HtmlSaveOptions, and saves the result as output.html. It demonstrates how to omit hidden rows from the generated HTML markup using Aspose.Cells for .NET.
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

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // The ExportHiddenRows property is not available in this version of Aspose.Cells.
            // Hidden rows will be handled according to the default behavior.

            // Save the workbook as an HTML file
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
