using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsAddInConversion
{
    public class AddInPdfConverter
    {
        /// <summary>
        /// Converts an XLSM workbook (which may contain Office Add‑Ins/macros UI) to PDF.
        /// The ConversionUtility.Convert method preserves the visual UI elements such as Ribbon XML.
        /// </summary>
        /// <param name="sourceXlsmPath">Full path to the source .xlsm file.</param>
        /// <param name="outputPdfPath">Full path where the resulting PDF will be saved.</param>
        public static void ConvertXlsmToPdf(string sourceXlsmPath, string outputPdfPath)
        {
            // Verify source file exists
            if (!File.Exists(sourceXlsmPath))
                throw new FileNotFoundException("Source XLSM file not found.", sourceXlsmPath);

            // Ensure output directory exists
            string? outDir = Path.GetDirectoryName(outputPdfPath);
            if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                Directory.CreateDirectory(outDir);

            // Perform conversion using Aspose.Cells utility
            ConversionUtility.Convert(sourceXlsmPath, outputPdfPath);
        }

        // Example usage
        public static void Run()
        {
            string sourceFile = @"C:\Samples\WorkbookWithAddIn.xlsm";
            string destFile   = @"C:\Samples\WorkbookWithAddIn.pdf";

            try
            {
                ConvertXlsmToPdf(sourceFile, destFile);
                Console.WriteLine($"Conversion completed: '{sourceFile}' → '{destFile}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public static class Program
    {
        public static void Main(string[] args)
        {
            // Optionally allow command‑line arguments for source and destination
            if (args.Length == 2)
            {
                try
                {
                    AddInPdfConverter.ConvertXlsmToPdf(args[0], args[1]);
                    Console.WriteLine($"Conversion completed: '{args[0]}' → '{args[1]}'");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during conversion: {ex.Message}");
                }
            }
            else
            {
                // Fallback to the example run method
                AddInPdfConverter.Run();
            }
        }
    }
}