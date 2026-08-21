// Title: Convert HTML to Excel in C# with Aspose.Cells – Preserve Currency, Percent & Numeric Formats
// Description: Step‑by‑step example that loads an HTML file into an Aspose.Cells Workbook using HtmlLoadOptions.KeepPrecision, then saves it as XLSX while keeping original currency symbols, percentages, and numeric precision.
// Keywords: Aspose.Cells | HTML to Excel conversion | C# | KeepPrecision | preserve currency format | preserve percentage format | numeric precision | HtmlLoadOptions | Excel export | financial data conversion
// Common Searches: Aspose.Cells keepprecision HTML to Excel C# | preserve currency symbols when converting HTML to XLSX | convert HTML percentages to Excel without losing format | load HTML into workbook with numeric precision | C# example HtmlLoadOptions KeepPrecision
// Developer Intent: Load an HTML document into an Aspose.Cells Workbook and export it as an Excel file while retaining the original numeric formatting such as currency symbols and percentage signs.
// Use Cases: Transform web‑based financial tables into Excel reports without losing monetary formatting. | Import scraped HTML data containing percentages for analysis in Excel. | Automate conversion of HTML invoices to XLSX while preserving exact currency values.
// AI Prompts: Show how to also keep date formats when converting HTML to Excel with Aspose.Cells. | Give an example of applying custom number formats after loading HTML with KeepPrecision. | Explain strategies for processing large HTML files efficiently using Aspose.Cells streaming.

using System;
using System.IO;
using Aspose.Cells;

// Step‑by‑step example that loads an HTML file into an Aspose.Cells Workbook using HtmlLoadOptions.KeepPrecision, then saves it as XLSX while keeping original currency symbols, percentages, and numeric precision.
class HtmlToExcelConverter
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlPath = "input.html";

        // Path for the resulting Excel file
        string excelPath = "output.xlsx";

        try
        {
            // Verify that the input HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Input file not found: {htmlPath}");
                return;
            }

            // Load the HTML file with options that preserve numeric precision and formatting
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                // KeepPrecision ensures that numeric strings (e.g., currency, percentages) are not
                // truncated or converted to less precise values during loading.
                KeepPrecision = true
            };

            // Load the HTML into a workbook
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Save the workbook as an Excel file (XLSX format)
            workbook.Save(excelPath, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion completed. Excel file saved to: {excelPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
        }
    }
}
