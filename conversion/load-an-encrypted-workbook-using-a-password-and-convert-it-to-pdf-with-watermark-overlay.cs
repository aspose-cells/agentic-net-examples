using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Path to the encrypted Excel file
        string inputPath = "encrypted.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        // Password required to open the workbook
        string password = "myPassword";

        try
        {
            // Load options with the password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };

            // Load the encrypted workbook
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Create a font for the watermark text
            RenderingFont font = new RenderingFont("Arial", 48)
            {
                Bold = true,
                Color = Color.Red
            };

            // Create a text watermark using the font
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
            {
                Rotation = 45,
                Opacity = 0.2f,
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                ScaleToPagePercent = 80
            };

            // Set PDF save options with the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as a PDF with the watermark applied
            string outputPath = "output.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (CellsException ex)
        {
            // Handle invalid password scenario
            if (ex.Message != null && ex.Message.IndexOf("invalid password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Failed to open workbook: Invalid password.");
            }
            else
            {
                Console.WriteLine($"CellsException: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}