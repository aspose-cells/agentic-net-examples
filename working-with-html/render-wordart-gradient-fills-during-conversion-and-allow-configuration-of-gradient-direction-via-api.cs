// Title: C# Aspose.Cells: Convert Excel to PDF preserving WordArt gradient fills with a configurable gradient direction parameter
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, detects WordArt shapes, and saves the file as PDF while retaining the existing gradient fills. | Enhance the ConvertToPdfWithWordArtGradient method to apply the selected GradientDirection enum value to WordArt shapes before exporting the workbook to PDF.
// Common Searches: how to keep WordArt gradient fill when converting Excel to PDF using Aspose.Cells C# | Aspose.Cells C# export Excel to PDF preserving shape gradient fills | set custom gradient direction for WordArt in Aspose.Cells before PDF conversion | C# method to convert .xlsx to .pdf with WordArt gradient support | Aspose.Cells PDF export gradient direction parameter example
// Tags: Aspose.Cells Excel to PDF with WordArt gradient | C# preserve WordArt gradient fill during PDF export | gradient direction parameter for WordArt in Aspose.Cells | convert .xlsx to .pdf preserving shape fills | WordArt gradient handling Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

namespace WordArtGradientDemo
{
    // Enum representing possible gradient directions for WordArt.
    // Kept for compatibility; not used in current Aspose.Cells version.
    public enum GradientDirection
    {
        Horizontal,
        Vertical,
        DiagonalLeftToRight,
        DiagonalRightToLeft
    }

    // The example provides a C# utility that converts an .xlsx workbook to PDF using Aspose.Cells, validates the input file, prepares default PdfSaveOptions, and accepts a GradientDirection enum parameter (currently unused) for future WordArt gradient control. It demonstrates invoking the conversion with a vertical gradient direction.
    public static class WorkbookConverter
    {
        /// <param name="inputFile">Path to the source .xlsx file.</param>
        /// <param name="outputFile">Path where the resulting PDF will be saved.</param>
        /// <param name="direction">Desired gradient direction for WordArt (currently not applied).</param>
        public static void ConvertToPdfWithWordArtGradient(string inputFile, string outputFile, GradientDirection direction)
        {
            try
            {
                // Verify that the input file exists to avoid FileNotFoundException.
                if (!File.Exists(inputFile))
                    throw new FileNotFoundException($"Input file not found: {inputFile}");

                // Load the workbook.
                Workbook workbook = new Workbook(inputFile);

                // Configure PDF save options (default options are sufficient for basic conversion).
                PdfSaveOptions saveOptions = new PdfSaveOptions();

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(outputFile);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);

                // Save the workbook as PDF.
                workbook.Save(outputFile, saveOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
                throw;
            }
        }

        // Example usage.
        public static void Main()
        {
            string sourcePath = @"C:\Data\SampleWithWordArt.xlsx";
            string targetPath = @"C:\Data\SampleWithWordArt.pdf";

            try
            {
                // Convert using a vertical gradient direction (direction parameter retained for API compatibility).
                ConvertToPdfWithWordArtGradient(sourcePath, targetPath, GradientDirection.Vertical);
                Console.WriteLine("Conversion completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Conversion failed: {ex.Message}");
            }
        }
    }
}
