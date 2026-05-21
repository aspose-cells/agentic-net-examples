using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class BatchExcelToPdf
{
    static void Main()
    {
        // Folder containing the source Excel files
        string inputFolder = "InputExcels";
        // Folder where the resulting PDF files will be saved
        string outputFolder = "OutputPdfs";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Verify the input directory exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder '{inputFolder}' does not exist. No files to process.");
            return;
        }

        // Process each .xlsx file in the input folder
        foreach (string sourcePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            // Extra safety: skip if the source file does not exist
            if (!File.Exists(sourcePath))
                continue;

            try
            {
                // Load options – load only data (if supported) to ignore charts/shapes
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
                // Uncomment the following line if your Aspose.Cells version supports it
                // loadOptions.LoadDataOnly = true;

                // PDF save options – each worksheet on a single page
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true
                };

                // Determine the destination PDF file path
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                string destPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                // Convert the Excel file to PDF
                ConversionUtility.Convert(sourcePath, loadOptions, destPath, pdfOptions);
                Console.WriteLine($"Converted '{sourcePath}' to '{destPath}'.");
            }
            catch (Exception ex)
            {
                // Log conversion errors without stopping the batch
                Console.WriteLine($"Error converting '{sourcePath}': {ex.Message}");
            }
        }
    }
}