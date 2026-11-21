// Title: Convert an Excel workbook to HTML with formula results using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export .xlsx to .html with calculated values | How to save Excel as HTML showing formula results instead of formulas in .NET | C# Aspose.Cells HtmlSaveOptions evaluate formulas before saving to HTML | Export Excel to HTML with formulas evaluated using Aspose.Cells library | Convert workbook to HTML and display formula outcomes in C#
// Tags: Aspose.Cells HtmlSaveOptions for HTML export | calculate workbook formulas before HTML save | save workbook as HTML with evaluated formulas | C# convert .xlsx to .html using Aspose.Cells | display formula results in HTML output

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, calls CalculateFormula() to evaluate all formulas, configures HtmlSaveOptions, and saves the workbook as an HTML file so that cells show the computed values rather than the original formula strings.
class ExcelToHtmlConverter
{
    static void Main()
    {
        // Paths for input Excel and output HTML files
        string excelPath = "input.xlsx";
        string htmlPath = "output.html";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Error: The input file \"{excelPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the Excel file
            Workbook workbook = new Workbook(excelPath);

            // Ensure all formulas are evaluated so their results are available
            workbook.CalculateFormula();

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // The ExportFormulasAsValues option is not required here because
            // CalculateFormula() already evaluates formulas before saving.

            // Save the workbook as an HTML file using the configured options
            workbook.Save(htmlPath, htmlOptions);

            Console.WriteLine($"Conversion successful. HTML saved to \"{htmlPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred during conversion: {ex.Message}");
        }
    }
}
