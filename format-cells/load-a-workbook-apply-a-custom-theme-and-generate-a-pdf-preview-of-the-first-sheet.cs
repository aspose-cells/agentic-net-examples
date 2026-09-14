// Title: Generate a PDF preview of the first worksheet from an Excel file with optional custom theme using Aspose.Cells for .NET
// AI Prompts: Create C# code that loads an .xlsx workbook with Aspose.Cells, checks for a .thmx theme file, and saves only the first sheet as a PDF preview. | Demonstrate how to hide all worksheets except the first one and configure PdfSaveOptions.OnePagePerSheet before exporting to PDF in Aspose.Cells. | Add robust error handling for missing workbook or theme files and for PDF saving failures when generating a preview with Aspose.Cells.
// Common Searches: aspnet c# export first sheet of excel to pdf using aspose.cells | aspose.cells hide other worksheets before saving pdf | pdfsaveoptions onepagepersheet example c# | check for custom .thmx theme file before converting excel to pdf with aspose.cells | handle missing input.xlsx error aspose.cells pdf export
// Tags: export first worksheet to PDF Aspose.Cells | hide worksheets except first Aspose.Cells | PdfSaveOptions OnePagePerSheet setting | custom .thmx theme handling Aspose.Cells | missing workbook file error handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The example loads 'input.xlsx' with Aspose.Cells, optionally detects a 'custom.thmx' theme (application not supported in older versions), hides all worksheets except the first, sets PdfSaveOptions.OnePagePerSheet, and saves the first sheet as 'preview.pdf'. It includes checks for missing files and comprehensive exception handling.
    class Program
    {
        static void Main()
        {
            try
            {
                string workbookPath = "input.xlsx";
                string themePath = "custom.thmx";
                string outputPath = "preview.pdf";

                // Verify that the input workbook exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Error: Workbook file not found: {workbookPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Apply custom theme if supported and the theme file exists
                if (File.Exists(themePath))
                {
                    try
                    {
                        // Aspose.Cells for .NET does not expose a direct SetTheme method in older versions.
                        // If a newer version with SetTheme is referenced, uncomment the line below:
                        // workbook.SetTheme(themePath);

                        // Placeholder for theme application – currently not supported.
                        Console.WriteLine("Custom theme file detected, but applying themes is not supported in this version.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Failed to apply theme. {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Warning: Theme file not found: {themePath}. Continuing without applying a custom theme.");
                }

                // Hide all worksheets except the first one to export only the first sheet
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    workbook.Worksheets[i].IsVisible = i == 0;
                }

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true
                };

                // Save the first sheet as a PDF preview
                try
                {
                    workbook.Save(outputPath, pdfOptions);
                    Console.WriteLine($"PDF preview saved successfully to: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving PDF: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
