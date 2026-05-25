using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsAddInPdfConversion
{
    class Program
    {
        static void Main()
        {
            string sourcePath = "input_with_addins.xlsx";
            string outputPath = "output_preserved.pdf";

            try
            {
                // Ensure the source workbook exists
                if (!File.Exists(sourcePath))
                {
                    // Create a minimal placeholder workbook if the file is missing
                    Workbook placeholder = new Workbook();
                    placeholder.Worksheets[0].Name = "Sheet1";
                    placeholder.Save(sourcePath);
                    Console.WriteLine($"Source file not found. Created placeholder workbook at '{sourcePath}'.");
                }

                // Convert the workbook to PDF while preserving interactive controls
                ConversionUtility.Convert(sourcePath, outputPath);
                Console.WriteLine($"Workbook converted to PDF successfully: {outputPath}");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}