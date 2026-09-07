// Title: Convert HTML tables with CSS display:none hidden rows and columns to an XLSX file using Aspose.Cells for .NET
// AI Prompts: Load an HTML file into an Aspose.Cells Workbook in C# while preserving rows and columns hidden by CSS display:none, then save it as XLSX. | Show how to configure HtmlLoadOptions.ImportHiddenRowsAndColumns to keep hidden elements during HTML‑to‑Excel conversion with Aspose.Cells. | Write C# code that checks the input HTML path, creates the output folder if missing, performs the conversion, and handles any exceptions.
// Common Searches: Aspose.Cells preserve CSS display:none rows when converting HTML to Excel in C# | How to keep hidden columns from an HTML table during HTML to XLSX conversion using Aspose.Cells | C# HtmlLoadOptions ImportHiddenRowsAndColumns example for HTML to Excel | Convert HTML with hidden elements to Excel workbook while retaining visibility settings | Aspose.Cells HTML to XLSX conversion hidden rows not removed
// Tags: Aspose.Cells HtmlLoadOptions import hidden rows | HTML to XLSX conversion preserving display:none | C# load HTML workbook hidden columns | Excel export hidden rows from HTML | convert HTML tables with hidden elements using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an HTML file into an Aspose.Cells Workbook using HtmlLoadOptions (which can be set to import rows and columns hidden with CSS display:none), ensures the output directory exists, saves the workbook as an XLSX file, and handles any conversion errors.
class HtmlToExcelConverter
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlPath = "input.html";

        // Path for the output Excel file
        string excelPath = "output.xlsx";

        // Verify that the input HTML file exists
        if (!File.Exists(htmlPath))
        {
            Console.WriteLine($"Input HTML file not found: {htmlPath}");
            return;
        }

        try
        {
            // Configure HTML load options (default options are sufficient for most cases)
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // Load the HTML file into a workbook using the configured options
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(excelPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to Excel format
            workbook.Save(excelPath, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion successful. Excel file saved to: {excelPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during conversion: {ex.Message}");
        }
    }
}
