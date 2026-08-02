// Title: Convert Excel to HTML with BestFit Layout (PresentationPreference) using Aspose.Cells for .NET
// Description: Shows how to export a workbook to HTML with Aspose.Cells by turning on PresentationPreference for automatic column‑width fitting and turning off full‑path links. The sample creates a workbook, adds data, configures HtmlSaveOptions, and saves the file.
// Keywords: Aspose.Cells | HTML conversion | PresentationPreference | BestFit layout | HtmlSaveOptions | C# example | disable full path links | Excel to HTML | responsive spreadsheet HTML | auto column width
// Common Searches: Aspose.Cells PresentationPreference HTML | C# save Excel as HTML best fit | disable full path links Aspose.Cells HTML output | convert workbook to HTML with auto column width | Aspose.Cells HtmlSaveOptions sample code
// Developer Intent: Export an Excel workbook to HTML with auto‑adjusted column widths (BestFit) and relative links using Aspose.Cells for .NET.
// Use Cases: Create web‑ready HTML views of spreadsheets where columns automatically fit their content. | Generate HTML reports for dashboards without absolute file paths, enabling easy embedding. | Produce responsive HTML files for multi‑page Excel documents while preserving layout consistency.
// AI Prompts: Write C# code that loads an existing .xlsx file and saves it as HTML with PresentationPreference enabled and IsFullPathLink set to false using Aspose.Cells. | Explain the impact of PresentationPreference on the HTML output and how to control column‑width behavior in Aspose.Cells. | Provide a step‑by‑step guide to convert each worksheet of a workbook into separate HTML files while keeping the BestFit layout.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlConversion
{
    // Shows how to export a workbook to HTML with Aspose.Cells by turning on PresentationPreference for automatic column‑width fitting and turning off full‑path links. The sample creates a workbook, adds data, configures HtmlSaveOptions, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Presentation Preference Demo");
            sheet.Cells["B1"].PutValue(123.45);
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // Set PresentationPreference to true for a more beautiful layout (BestFit)
            htmlOptions.PresentationPreference = true;
            // Optional: avoid using full path links in generated HTML files
            htmlOptions.IsFullPathLink = false;

            // Save the workbook as HTML using the configured options
            workbook.Save("PresentationPreference.html", htmlOptions);

            Console.WriteLine("Workbook successfully saved as HTML with PresentationPreference enabled.");
        }
    }
}
