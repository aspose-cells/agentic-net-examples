using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel workbook (XLSX)
            string sourcePath = "input.xlsx";

            // Path for the output CSV file
            string destPath = "output.csv";

            try
            {
                // Load options for the source file (XLSX format)
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

                // Save options for CSV with separators kept for blank rows
                TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    KeepSeparatorsForBlankRow = true, // Preserve separators on empty rows
                    Separator = ',',                  // Use comma as the delimiter
                    Encoding = System.Text.Encoding.UTF8
                };

                // Perform the conversion using Aspose.Cells ConversionUtility
                ConversionUtility.Convert(sourcePath, loadOptions, destPath, saveOptions);

                Console.WriteLine($"Conversion completed successfully: '{sourcePath}' -> '{destPath}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}