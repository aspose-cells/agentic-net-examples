// Title: Export each Excel chart to a separate PDF file with per‑chart CultureInfo using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loops through all worksheets in a Workbook, accesses each Chart object, assigns a specific CultureInfo to workbook.Settings.CultureInfo, and saves the chart as an individual PDF file. | Show how to construct PDF filenames that embed the chart index and the applied locale (e.g., Chart_1_en-US.pdf) while exporting charts with Aspose.Cells. | Provide robust error handling for missing Excel files and for exceptions thrown during chart‑to‑PDF conversion in a .NET console application.
// Common Searches: how to export Excel charts to separate PDF files with different locales using Aspose.Cells C# | Aspose.Cells set CultureInfo for each chart before saving as PDF | C# iterate workbook charts and generate PDF per chart with language‑specific formatting | export chart to pdf with custom globalization settings Aspose.Cells .NET | save individual chart PDFs from an Excel workbook using locale‑aware settings
// Tags: export chart to pdf Aspose.Cells | per‑chart CultureInfo setting .NET | chart globalization Aspose.Cells | C# workbook chart PDF conversion | locale‑aware chart export Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, defines a set of culture identifiers, and iterates through every worksheet and its charts. For each chart it selects a culture, applies it to workbook.Settings.CultureInfo to affect rendering, and exports the chart to a PDF file named with the chart number and culture code. The program reports the total number of exported charts and includes error handling for missing files and conversion failures.
class ExportChartsToPdf
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Example list of culture names to apply per chart
        string[] cultureNames = new[] { "en-US", "fr-FR", "de-DE", "ja-JP" };
        int chartCounter = 0;

        try
        {
            // Iterate through all worksheets and their charts
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    // Select a culture for the current chart
                    string cultureName = cultureNames[chartCounter % cultureNames.Length];
                    CultureInfo cultureInfo = new CultureInfo(cultureName);

                    // Apply the culture to the workbook (affects chart rendering)
                    workbook.Settings.CultureInfo = cultureInfo;

                    // Export the chart to a PDF file
                    string pdfFileName = $"Chart_{chartCounter + 1}_{cultureName}.pdf";
                    chart.ToPdf(pdfFileName);

                    chartCounter++;
                }
            }

            Console.WriteLine($"{chartCounter} chart(s) exported to PDF with custom globalization settings.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during chart export: {ex.Message}");
        }
    }
}
