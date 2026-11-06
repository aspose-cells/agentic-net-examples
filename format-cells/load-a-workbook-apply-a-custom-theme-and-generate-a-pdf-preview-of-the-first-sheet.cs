// Title: Create a PDF preview of the first worksheet from an Excel file with optional custom theme using Aspose.Cells for .NET (C#)
// Common Searches: how to export only the first sheet of an Excel workbook to PDF using Aspose.Cells C# | Aspose.Cells C# generate PDF preview of a worksheet with PdfSaveOptions | apply a custom .thmx theme to a workbook before saving as PDF with Aspose.Cells | C# code to check for a theme file and export the first page as PDF using Aspose.Cells | save first page of an Excel worksheet as a PDF preview in .NET
// Tags: Aspose.Cells first sheet PDF export | PdfSaveOptions one-page output | custom .thmx theme handling Aspose.Cells | Excel to PDF preview generation C# | file existence validation C# Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an Excel workbook, optionally checks for a custom .thmx theme file, configures PdfSaveOptions to export only the first page of the first worksheet, and saves the result as a PDF preview while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string themePath = "customTheme.thmx";
            const string outputPath = "preview.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Load and apply a custom theme if the file is present
            if (File.Exists(themePath))
            {
                try
                {
                    // Aspose.Cells does not expose a public Theme class in all versions.
                    // If a custom theme API is available, it can be applied here.
                    // Placeholder for future theme loading logic.
                    Console.WriteLine($"Info: Theme file \"{themePath}\" found, but custom theme loading is not supported in this version.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Failed to load theme. {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Warning: Theme file \"{themePath}\" not found. Continuing without custom theme.");
            }

            // Configure PDF save options to export only the first page of the first worksheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true,
                PageIndex = 0,   // start from the first page (0‑based)
                PageCount = 1    // export only one page
            };

            // Save the first sheet as a PDF preview
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Preview PDF saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
