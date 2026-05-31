using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Saving;

namespace AsposeCellsAddInPdfConversion
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths for the source CSV file and the output PDF file
                string csvPath = "data.csv";
                string pdfPath = "result.pdf";

                // -----------------------------------------------------------------
                // 1. Create a sample CSV file (in real scenarios the file already exists)
                // -----------------------------------------------------------------
                File.WriteAllText(csvPath,
                    "Product,Price,Quantity\n" +
                    "Apple,1.5,100\n" +
                    "Banana,0.75,150\n" +
                    "Orange,1.25,120");

                // Ensure the CSV file exists before proceeding
                if (!File.Exists(csvPath))
                    throw new FileNotFoundException("CSV source file not found.", csvPath);

                // -----------------------------------------------------------------
                // 2. Load the CSV file into a workbook.
                //    LoadOptions with LoadFormat.Csv tells Aspose.Cells to treat the file as CSV.
                // -----------------------------------------------------------------
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
                Workbook workbook = new Workbook(); // empty workbook
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells.ImportCSV(csvPath, ",", true, 0, 0);

                // -----------------------------------------------------------------
                // 3. Save the workbook as PDF.
                //    PdfSaveOptions can be used to fine‑tune the PDF output.
                // -----------------------------------------------------------------
                PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
                // Example: embed all fonts (property may not be available in older versions)
                // pdfSaveOptions.EmbedStandardFonts = true; // Uncomment if supported

                // Use ConversionUtility to apply both load and save options
                ConversionUtility.Convert(csvPath, loadOptions, pdfPath, pdfSaveOptions);

                // -----------------------------------------------------------------
                // 4. Verify the conversion result.
                // -----------------------------------------------------------------
                Console.WriteLine(File.Exists(pdfPath)
                    ? $"CSV file successfully converted to PDF: {pdfPath}"
                    : "PDF conversion failed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                // Clean up temporary CSV file (optional)
                string csvPath = "data.csv";
                if (File.Exists(csvPath))
                {
                    try { File.Delete(csvPath); } catch { /* ignore cleanup errors */ }
                }
            }
        }
    }
}