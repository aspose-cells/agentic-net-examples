using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsPdfToJson
{
    class Program
    {
        static void Main()
        {
            // Paths for source PDF, temporary Excel, and final JSON output
            string pdfPath = "input.pdf";
            string tempExcelPath = "temp.xlsx";
            string jsonOutputPath = "output.json";

            try
            {
                // 1. Convert PDF (treated as an Excel file) to an intermediate XLSX file
                ConversionUtility.Convert(pdfPath, tempExcelPath);

                // 2. Load the intermediate workbook
                Workbook workbook = new Workbook(tempExcelPath);

                // 3. Configure JSON save options
                JsonSaveOptions jsonOptions = new JsonSaveOptions
                {
                    // Export as a JSON object even if there is only one worksheet
                    AlwaysExportAsJsonObject = true,
                    // Assume the first row contains headers (optional, adjust as needed)
                    HasHeaderRow = true,
                    // Export empty cells as null (optional)
                    ExportEmptyCells = true,
                    // Indent the JSON for readability (optional)
                    Indent = "  "
                };

                // 4. Save the workbook as JSON
                workbook.Save(jsonOutputPath, jsonOptions);

                Console.WriteLine($"Conversion completed successfully. JSON saved to '{jsonOutputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
            finally
            {
                // Clean up the temporary Excel file if it exists
                if (File.Exists(tempExcelPath))
                {
                    try
                    {
                        File.Delete(tempExcelPath);
                    }
                    catch
                    {
                        // Ignored – if deletion fails, the file will remain on disk
                    }
                }
            }
        }
    }
}