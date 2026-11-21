// Title: Convert an HTML file to XLSX and enable automatic formula calculation with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells convert HTML to Excel workbook with formulas evaluated automatically | How to load an HTML file into Aspose.Cells and set calculation mode to automatic | Aspose.Cells calculate formulas after importing HTML in a .NET application | Saving HTML as .xlsx using Aspose.Cells while preserving formula results | C# code to force formula recalculation when converting HTML to XLSX with Aspose.Cells
// Tags: Aspose.Cells HTML import to XLSX | automatic calculation mode workbook Aspose.Cells | calculate formulas after HTML load Aspose.Cells | HtmlLoadOptions for HTML to workbook conversion | save workbook as Xlsx using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for an input HTML file, loads it into an Aspose.Cells Workbook with HtmlLoadOptions, forces formula evaluation via CalculateFormula, and saves the result as an XLSX file with automatic calculation mode enabled.
class Program
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlPath = "input.html";

        // Path to the output Excel file
        string excelPath = "output.xlsx";

        // Verify that the input HTML file exists
        if (!File.Exists(htmlPath))
        {
            Console.WriteLine($"Error: The input file \"{htmlPath}\" was not found.");
            return;
        }

        try
        {
            // Load the HTML file into a workbook
            Workbook workbook = new Workbook(htmlPath, new HtmlLoadOptions());

            // Force calculation of any formulas present in the workbook
            workbook.CalculateFormula();

            // Save the workbook as an Excel file
            workbook.Save(excelPath, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion successful. Excel file saved to \"{excelPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during conversion
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
