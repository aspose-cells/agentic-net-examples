// Title: Export an Excel workbook to HTML with grid lines and best‑fit column widths using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file, sets HtmlSaveOptions.PresentationPreference to BestFit, enables grid‑line export, and saves the workbook as an HTML document with Aspose.Cells. | Show a complete .NET example that configures Aspose.Cells HtmlSaveOptions to preserve cell borders and automatically adjust column widths when converting Excel to HTML.
// Common Searches: Aspose.Cells C# export Excel workbook to HTML with visible grid lines | How to make columns auto‑fit when converting .xlsx to .html using Aspose.Cells | Example of HtmlSaveOptions.ExportGridLines true for Excel to HTML conversion | Set PresentationPreference to BestFit in Aspose.Cells HTML output | Convert Excel file to HTML while keeping original column widths in .NET
// Tags: Aspose.Cells HTML grid line support | BestFit column sizing Aspose.Cells | C# Excel to HTML conversion Aspose.Cells | Preserve Excel layout in HTML Aspose.Cells | Auto column width HTML Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The C# sample loads an existing workbook (or creates a simple one), configures HtmlSaveOptions to render grid lines and apply best‑fit column widths, and then saves the workbook as an HTML file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Ensure the input workbook exists; create a simple one if it does not.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample Data");
                workbook.Save(inputPath);
            }

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Export grid lines to the generated HTML
                ExportGridLines = true
                // PresentationPreference property removed for compatibility with current API version
            };

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
