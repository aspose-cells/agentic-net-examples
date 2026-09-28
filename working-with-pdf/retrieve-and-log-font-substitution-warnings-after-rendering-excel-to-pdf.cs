// Title: Log font substitution warnings when converting an Excel workbook to PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that saves an Excel workbook as PDF with Aspose.Cells and writes any font substitution warnings to the console. | Show how to use .NET reflection to read the FontSubstitutionWarnings collection from a Workbook after PDF export. | Create a reusable method that extracts original and substituted font names from Aspose.Cells PDF conversion warnings without relying on a specific library version.
// Common Searches: Aspose.Cells retrieve font substitution warnings after converting Excel to PDF in C# | Log missing or substituted fonts during Excel to PDF conversion using Aspose.Cells .NET | Access FontSubstitutionWarnings property via reflection for different Aspose.Cells versions | How to get list of substituted fonts when saving workbook as PDF with Aspose.Cells
// Tags: Aspose.Cells PDF export font substitution warnings | C# reflection access FontSubstitutionWarnings | Excel to PDF missing fonts handling Aspose.Cells | version‑agnostic warning retrieval Aspose.Cells | log original substituted fonts Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel workbook, saves it as a PDF with Aspose.Cells, then uses reflection to read the FontSubstitutionWarnings collection and logs each original and substituted font name.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Set PDF save options (customize as needed)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Render the workbook to PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");

            // Retrieve font substitution warnings using reflection to stay compatible with different library versions
            try
            {
                var warningsProp = workbook.GetType().GetProperty("FontSubstitutionWarnings");
                if (warningsProp != null)
                {
                    var warnings = warningsProp.GetValue(workbook) as System.Collections.IEnumerable;
                    if (warnings != null)
                    {
                        foreach (var warning in warnings)
                        {
                            var fontNameProp = warning.GetType().GetProperty("FontName");
                            var subFontProp = warning.GetType().GetProperty("SubstitutedFontName");
                            string originalFont = fontNameProp?.GetValue(warning)?.ToString() ?? "N/A";
                            string substitutedFont = subFontProp?.GetValue(warning)?.ToString() ?? "N/A";
                            Console.WriteLine($"Original Font: {originalFont}, Substituted Font: {substitutedFont}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning retrieval failed: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
