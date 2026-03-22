using System;
using Aspose.Cells;

namespace TimelineConversion
{
    public class Converter
    {
        /// <summary>
        /// Converts a timeline spreadsheet from XLS to XLSX format while preserving all data and layout.
        /// </summary>
        /// <param name="sourcePath">Full path of the source .xls file.</param>
        /// <param name="destPath">Full path where the .xlsx file will be saved.</param>
        public static void ConvertXlsToXlsx(string sourcePath, string destPath)
        {
            try
            {
                // If the source file does not exist, create a simple XLS file for demonstration.
                if (!System.IO.File.Exists(sourcePath))
                {
                    var sampleWb = new Workbook();
                    sampleWb.Worksheets[0].Name = "Timeline";
                    sampleWb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                    // Save without specifying format; format is inferred from the file extension.
                    sampleWb.Save(sourcePath);
                }

                // Load the XLS workbook and save it as XLSX.
                var workbook = new Workbook(sourcePath);
                // Save without specifying format; format is inferred from the file extension.
                workbook.Save(destPath);

                Console.WriteLine($"Conversion successful: \"{sourcePath}\" → \"{destPath}\"");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
                throw;
            }
        }

        // Example usage
        public static void Main()
        {
            // Specify the input XLS file and the desired output XLSX file.
            string sourceFile = @"C:\Data\TimelineReport.xls";
            string outputFile = @"C:\Data\TimelineReport.xlsx";

            ConvertXlsToXlsx(sourceFile, outputFile);
        }
    }
}