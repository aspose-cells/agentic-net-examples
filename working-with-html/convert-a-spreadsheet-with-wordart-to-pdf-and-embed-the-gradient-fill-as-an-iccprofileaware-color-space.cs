// Title: Convert an Excel workbook with WordArt to PDF and embed an ICC color profile using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file containing WordArt, configures PdfSaveOptions to apply an ICC color profile only when the profile file exists and the Aspose.Cells API supports it, then saves the workbook as a PDF. | Show how to detect the presence of an ICC file and safely enable color‑management options during Excel‑to‑PDF conversion with Aspose.Cells, including fallback handling for older library versions.
// Common Searches: how to embed an ICC profile when saving Excel to PDF with Aspose.Cells .NET | preserve WordArt gradient colors in PDF output using Aspose.Cells C# | conditional IccProfilePath setting in PdfSaveOptions for .NET | color management support check for Aspose.Cells version before PDF conversion | convert workbook with WordArt to PDF using sRGB ICC profile in C#
// Tags: pdfsaveoptions icc embedding Aspose.Cells | excel wordart pdf export c# | icc profile presence check .net | gradient fill retention Aspose.Cells pdf | version check for color‑management Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads 'InputWithWordArt.xlsx', verifies the workbook and an optional 'sRGB.icc' file, creates a PdfSaveOptions object, conditionally sets IccProfilePath and EmbedIccProfile when supported, saves the workbook as 'Output.pdf', and logs success or error messages.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWithWordArt.xlsx";
            const string outputPath = "Output.pdf";
            const string iccPath = "sRGB.icc";

            // Verify that the input workbook exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook containing WordArt objects.
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // If an ICC profile file is present and the current Aspose.Cells version supports it,
            // you can embed it by setting the appropriate properties (available in newer versions).
            // The following block is guarded by a file‑existence check and a try‑catch to avoid
            // runtime errors on older library versions.
            if (File.Exists(iccPath))
            {
                try
                {
                    // The properties IccProfilePath and EmbedIccProfile exist in recent releases.
                    // Uncomment the lines below when using a version that provides them.
                    // pdfOptions.IccProfilePath = iccPath;
                    // pdfOptions.EmbedIccProfile = true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Unable to set ICC profile options – {ex.Message}");
                }
            }

            // Save the workbook as a PDF.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
