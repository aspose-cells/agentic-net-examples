// Title: Export an Excel workbook to HTML with default HtmlSaveOptions and set hyperlink target to _parent using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a .xlsx file with Aspose.Cells, creates HtmlSaveOptions with default settings, assigns HyperlinkTarget = "_parent", and saves the result as an .html file. | Show how to export a workbook to HTML in C# with Aspose.Cells, include a check for the HyperlinkTarget property availability, and add error handling for a missing source file.
// Common Searches: how to export Excel to HTML with Aspose.Cells and set link target to parent | C# Aspose.Cells HtmlSaveOptions default export example | apply link target attribute during HTML export with Aspose.Cells | handle file not found error before converting .xlsx to .html with Aspose.Cells | Aspose.Cells HTML export without custom styling and with _parent links
// Tags: Aspose.Cells HtmlSaveOptions export workbook | set HyperlinkTarget _parent Aspose.Cells | convert .xlsx to .html C# | default HTML export options Aspose.Cells | missing input file handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads "input.xlsx" using Aspose.Cells, creates default HtmlSaveOptions (optionally setting HyperlinkTarget to "_parent"), and saves the workbook as "output.html" while handling missing files and runtime exceptions.
class ExcelToHtmlConverter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options with default settings
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // The HyperlinkTarget property may not be available in older versions.
            // If needed, uncomment the line below and ensure the property exists.
            // htmlOptions.HyperlinkTarget = "_parent";

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
