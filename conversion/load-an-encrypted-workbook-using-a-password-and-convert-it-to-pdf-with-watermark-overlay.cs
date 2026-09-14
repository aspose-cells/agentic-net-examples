// Title: Load a password‑protected Excel workbook and convert it to a PDF with a diagonal semi‑transparent watermark using Aspose.Cells for .NET
// AI Prompts: Load a password‑protected .xlsx workbook, define a red 48‑pt Arial watermark rotated 45° with 20% opacity, and export the file to PDF using Aspose.Cells. | Use LoadOptions to supply the workbook password, create a RenderingWatermark (font, rotation, opacity, alignment), assign it to PdfSaveOptions, and save the protected workbook as a PDF in C#.
// Common Searches: Aspose.Cells .NET how to provide password when loading an encrypted Excel file | C# convert password‑protected Excel workbook to PDF with diagonal watermark using Aspose.Cells | Set watermark opacity and rotation in PdfSaveOptions Aspose.Cells | Example of LoadOptions.Password property for encrypted workbook Aspose.Cells
// Tags: load encrypted workbook Aspose.Cells .NET | pdf conversion with diagonal watermark Aspose.Cells | renderingwatermark configuration Aspose.Cells | password protected Excel to PDF Aspose.Cells | pdfsaveoptions watermark Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample creates (if needed) an encrypted Excel file, loads it with LoadOptions using the supplied password, builds a red 48‑pt Arial watermark rotated 45° at 20% opacity, assigns the watermark to PdfSaveOptions, and saves the workbook as a PDF with the watermark applied.
class Program
{
    static void Main()
    {
        // Path to the encrypted Excel file
        string inputFile = "encrypted.xlsx";

        // Password required to open the workbook
        string password = "myPassword";

        // Ensure the input file exists; if not, create a simple encrypted workbook for demo purposes
        if (!File.Exists(inputFile))
        {
            try
            {
                Workbook tempWb = new Workbook();
                tempWb.Worksheets[0].Cells["A1"].PutValue("Sample data");
                // Set password to encrypt the workbook
                tempWb.Settings.Password = password;
                tempWb.Save(inputFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create placeholder workbook: {ex.Message}");
                return;
            }
        }

        Workbook? workbook = null;
        try
        {
            // Load the workbook with the password using LoadOptions
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };
            workbook = new Workbook(inputFile, loadOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Create a font for the watermark text
        RenderingFont watermarkFont = new RenderingFont("Arial", 48)
        {
            Bold = true,
            Color = Color.Red
        };

        // Create a text watermark
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
        {
            Rotation = 45,                 // Rotate watermark
            Opacity = 0.2f,                // Semi‑transparent
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,
            ScaleToPagePercent = 80        // Scale relative to page size
        };

        // Set PDF save options and assign the watermark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as PDF with the watermark applied
        string outputFile = "output.pdf";
        try
        {
            workbook?.Save(outputFile, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving PDF: {ex.Message}");
        }
    }
}
