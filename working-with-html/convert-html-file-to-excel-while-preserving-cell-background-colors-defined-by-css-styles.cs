// Title: Convert HTML to Excel (XLSX) in C# with Aspose.Cells while preserving CSS background colors
// AI Prompts: Load an HTML file into an Aspose.Cells Workbook using C# and save it as XLSX, keeping all CSS background colors intact. | Show how to set up HtmlLoadOptions in Aspose.Cells to retain cell styling when converting a styled HTML table to Excel. | Implement robust error handling that checks for a missing HTML source file and creates the output directory before exporting the workbook with preserved colors.
// Common Searches: c# aspose.cells convert html table to xlsx preserving cell background colors | how to keep css styles when exporting html to excel using asp.net | asp.net load html with aspose.cells and retain background color formatting | html to excel conversion with aspose.cells preserving styling in .net core
// Tags: Aspose.Cells HTML to XLSX conversion | preserve CSS background colors in Excel export | C# HtmlLoadOptions styling retention | styled HTML table to Excel workbook | error handling for missing HTML input Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an HTML file into an Aspose.Cells Workbook with default HtmlLoadOptions, verifies the input, creates the output folder if needed, and saves the workbook as an XLSX file, automatically retaining any CSS‑defined cell background colors.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Verify that the input HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: Input file not found: {htmlPath}");
                return;
            }

            // Configure load options for HTML (no specific formatting option needed)
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // Load the HTML file into a workbook
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Path for the resulting Excel file
            string excelPath = "output.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(excelPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an Excel file (XLSX format)
            workbook.Save(excelPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to: {excelPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
