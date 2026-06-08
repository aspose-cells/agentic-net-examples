using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPasswordCheck
{
    public static class ExcelProtectionHelper
    {
        /// <summary>
        /// Determines whether the specified Excel file is password protected (encrypted).
        /// </summary>
        /// <param name="filePath">Full path to the Excel file.</param>
        /// <returns>True if the file requires a password to open; otherwise, false.</returns>
        public static bool IsFilePasswordProtected(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return false;
                }

                // Detect the file format and retrieve its metadata.
                FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);

                // The IsEncrypted property indicates whether the document is encrypted.
                return fileInfo.IsEncrypted;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error checking protection for '{filePath}': {ex.Message}");
                return false;
            }
        }

        // Example usage
        public static void RunDemo()
        {
            string unprotectedPath = "sample.xlsx";
            string protectedPath = "protected.xlsx";

            Console.WriteLine($"'{unprotectedPath}' protected? {IsFilePasswordProtected(unprotectedPath)}");
            Console.WriteLine($"'{protectedPath}' protected? {IsFilePasswordProtected(protectedPath)}");
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                ExcelProtectionHelper.RunDemo();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}