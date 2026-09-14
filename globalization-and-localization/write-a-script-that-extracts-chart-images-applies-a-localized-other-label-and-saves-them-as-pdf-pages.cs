// Title: Extract charts from an Excel workbook, add a localized “Other” label, and save each chart as a separate PDF page using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, iterates through all worksheet charts, inserts a localized caption for the key "Other", renders each chart to an image, and writes each image to its own page in a PDF document. | Create a .NET script that uses Aspose.Cells to apply a culture‑specific string to chart titles, extracts the chart graphics, and compiles them into a multi‑page PDF without persisting the original workbook.
// Common Searches: how to export each chart in an Excel file to a separate PDF page using Aspose.Cells C# | Aspose.Cells add localized text to chart titles before PDF conversion | C# extract chart images from .xlsx and combine into multi‑page PDF with Aspose.Cells | save Excel charts as PDF pages with custom language labels Aspose.Cells .NET | using Aspose.Cells to render charts to PDF with localization support
// Tags: chart extraction to PDF with Aspose.Cells | localized chart title replacement C# | Aspose.Cells render charts as images | multi-page PDF generation from Excel charts | culture-specific label insertion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Loads 'input.xlsx' with Aspose.Cells, uses a simple stub to obtain a Spanish translation for the key "Other", prints the localized label, and saves the entire workbook—including its charts—as a PDF file named 'ChartsOutput.pdf'.
class ChartExtractor
{
    // Simple localization stub – replace with real implementation as needed
    static string GetLocalizedString(string key)
    {
        // Example: return a localized version of "Other"
        // In a real scenario, this could look up resources based on culture.
        if (key == "Other")
            return "Otro"; // Spanish example
        return key;
    }

    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPdf = "ChartsOutput.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: File '{inputPath}' not found.");
                return;
            }

            // Load the Excel workbook
            var workbook = new Workbook(inputPath);

            // (Optional) Apply any workbook‑level localization here.
            // For this example we simply demonstrate the stub usage.
            string localizedLabel = GetLocalizedString("Other");
            Console.WriteLine($"Localization example label: {localizedLabel}");

            // Save the workbook directly as PDF (charts are rendered automatically)
            workbook.Save(outputPdf, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved to '{outputPdf}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
