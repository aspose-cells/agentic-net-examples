using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Rendering;
using Aspose.Cells.Saving;

class BatchXlsxToPdfA1a
{
    static void Main()
    {
        // Directory containing the source XLSX files
        string sourceDirectory = @"C:\InputXlsx";
        // Directory where the converted PDF/A‑1a files will be saved
        string outputDirectory = @"C:\OutputPdfA1a";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputDirectory);

        // Get all .xlsx files in the source directory (non‑recursive)
        string[] xlsxFiles = Directory.GetFiles(sourceDirectory, "*.xlsx", SearchOption.TopDirectoryOnly);

        // Parallel conversion of each file
        Parallel.ForEach(xlsxFiles, sourcePath =>
        {
            try
            {
                // Determine the output PDF file name (same base name, .pdf extension)
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                string destPath = Path.Combine(outputDirectory, fileNameWithoutExt + ".pdf");

                // Load options – specify that the source is an XLSX workbook
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

                // Save options – configure PDF/A‑1a compliance
                PdfSaveOptions saveOptions = new PdfSaveOptions
                {
                    // Set the PDF compliance level to PDF/A‑1a
                    Compliance = PdfCompliance.PdfA1a
                };

                // Perform the conversion using Aspose.Cells ConversionUtility
                ConversionUtility.Convert(sourcePath, loadOptions, destPath, saveOptions);

                Console.WriteLine($"Converted: {sourcePath} -> {destPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting '{sourcePath}': {ex.Message}");
            }
        });

        Console.WriteLine("Batch conversion completed.");
    }
}