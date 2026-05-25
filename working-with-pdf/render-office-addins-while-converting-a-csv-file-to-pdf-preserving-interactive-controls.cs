using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsAddInConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Paths for source CSV and destination PDF
            string csvPath = "sample.csv";
            string pdfPath = "output.pdf";

            // ------------------------------------------------------------
            // 1. Create a sample CSV file (simulating an Office Add‑In export)
            // ------------------------------------------------------------
            File.WriteAllText(csvPath,
                "Product,Price,Quantity,Link\n" +
                "Apple,1.20,100,\"https://example.com/apple\"\n" +
                "Banana,0.80,150,\"https://example.com/banana\"\n" +
                "Cherry,2.00,75,\"https://example.com/cherry\"");

            try
            {
                // ------------------------------------------------------------
                // 2. Verify the CSV file exists before conversion
                // ------------------------------------------------------------
                if (!File.Exists(csvPath))
                    throw new FileNotFoundException("CSV source file not found.", csvPath);

                // ------------------------------------------------------------
                // 3. Prepare load options for CSV format
                // ------------------------------------------------------------
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);

                // ------------------------------------------------------------
                // 4. Prepare PDF save options
                // ------------------------------------------------------------
                PdfSaveOptions saveOptions = new PdfSaveOptions
                {
                    // Keep interactive elements such as hyperlinks
                    CheckExcelRestriction = false
                };

                // ------------------------------------------------------------
                // 5. Convert CSV to PDF using the ConversionUtility rule
                // ------------------------------------------------------------
                ConversionUtility.Convert(csvPath, loadOptions, pdfPath, saveOptions);

                Console.WriteLine($"CSV file successfully converted to PDF: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Conversion failed: {ex.Message}");
            }
            finally
            {
                // Clean up temporary CSV file (optional)
                if (File.Exists(csvPath))
                {
                    try
                    {
                        File.Delete(csvPath);
                    }
                    catch
                    {
                        // Ignored – cleanup failure should not crash the app
                    }
                }
            }
        }
    }
}