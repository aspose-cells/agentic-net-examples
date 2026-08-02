// Title: Convert HTML to Excel with Currency, Percentage & Numeric Formatting – Aspose.Cells C#
// Description: Loads an HTML file using HtmlLoadOptions.KeepPrecision, converts numeric strings (including currency and percentages) to true numeric cells with ConvertStringToNumericValue, and saves the workbook as XLSX while preserving original numeric formats.
// Keywords: Aspose.Cells HTML to Excel | C# convert HTML to XLSX | preserve currency formatting | percentage numeric conversion | HtmlLoadOptions KeepPrecision | ConvertStringToNumericValue | Excel numeric precision | HTML financial report to Excel
// Common Searches: Aspose.Cells keep currency formatting when converting HTML to XLSX | C# convert HTML percentages to numeric cells in Excel | HtmlLoadOptions KeepPrecision example | ConvertStringToNumericValue after HTML import | HTML to Excel conversion preserving numeric formats
// Developer Intent: Generate an Excel workbook from an HTML file while retaining original numeric, currency, and percentage formats.
// Use Cases: Transform a web‑based financial statement saved as HTML into an Excel file where monetary values stay numeric for further calculations. | Export a KPI dashboard HTML page to Excel, keeping percentage metrics as numeric cells for charting. | Batch‑process HTML invoices, converting them to XLSX while preserving exact amounts and tax rates.
// AI Prompts: Write C# code that uses Aspose.Cells to load an HTML file with KeepPrecision, converts string cells to numeric values, and saves as XLSX. | Explain the effect of ConvertStringToNumericValue on currency and percentage cells after loading HTML with Aspose.Cells. | Suggest additional HtmlLoadOptions settings that improve numeric format retention during HTML‑to‑Excel conversion.

using System;
using System.IO;
using Aspose.Cells;

// Loads an HTML file using HtmlLoadOptions.KeepPrecision, converts numeric strings (including currency and percentages) to true numeric cells with ConvertStringToNumericValue, and saves the workbook as XLSX while preserving original numeric formats.
class HtmlToExcelConverter
{
    static void Main()
    {
        // Paths to the source HTML file and the destination Excel file
        string htmlFilePath = "input.html";
        string excelFilePath = "output.xlsx";

        // Verify that the input HTML file exists
        if (!File.Exists(htmlFilePath))
        {
            Console.WriteLine($"Error: The HTML file \"{htmlFilePath}\" was not found.");
            return;
        }

        try
        {
            // Load the HTML file with options that keep numeric precision.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                KeepPrecision = true // Preserve the original numeric precision when possible.
            };

            // Create a Workbook instance by loading the HTML file.
            Workbook workbook = new Workbook(htmlFilePath, loadOptions);

            // Convert any string values that can be interpreted as numbers (e.g., "123", "45%")
            // to actual numeric types to retain numeric formatting such as currency and percentages.
            workbook.Worksheets[0].Cells.ConvertStringToNumericValue();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(excelFilePath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an Excel file (XLSX format).
            workbook.Save(excelFilePath, SaveFormat.Xlsx);

            Console.WriteLine("HTML has been successfully converted to Excel with numeric formats preserved.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during conversion: {ex.Message}");
        }
    }
}
