using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class VerifyExtractedFileExtension
    {
        /// <summary>
        /// Verifies that the extension of the extracted file matches the original file's format.
        /// </summary>
        /// <param name="originalFilePath">Path to the original file.</param>
        /// <param name="extractedFilePath">Path to the extracted file to be verified.</param>
        public static void Run(string originalFilePath, string extractedFilePath)
        {
            try
            {
                // Ensure both files exist to avoid FileNotFoundException
                if (!File.Exists(originalFilePath))
                {
                    Console.WriteLine($"Original file not found: {originalFilePath}");
                    return;
                }

                if (!File.Exists(extractedFilePath))
                {
                    Console.WriteLine($"Extracted file not found: {extractedFilePath}");
                    return;
                }

                // Get the original file extension (e.g., ".xlsx")
                string originalExtension = Path.GetExtension(originalFilePath).ToLowerInvariant();

                // Detect the format of the extracted file using Aspose.Cells utility
                FileFormatInfo extractedInfo = FileFormatUtil.DetectFileFormat(extractedFilePath);

                // Convert the detected FileFormatType to a SaveFormat enum
                SaveFormat detectedSaveFormat = FileFormatUtil.FileFormatToSaveFormat(extractedInfo.FileFormatType);

                // Convert the SaveFormat to a file extension (e.g., ".xlsx")
                string detectedExtension = FileFormatUtil.SaveFormatToExtension(detectedSaveFormat).ToLowerInvariant();

                // Compare the original extension with the detected extension
                bool isMatch = string.Equals(originalExtension, detectedExtension, StringComparison.OrdinalIgnoreCase);

                // Output the verification result
                Console.WriteLine($"Original file extension : {originalExtension}");
                Console.WriteLine($"Detected file extension : {detectedExtension}");
                Console.WriteLine($"Extension match          : {isMatch}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    internal class Program
    {
        // Entry point required for console application
        private static void Main(string[] args)
        {
            // Expect two arguments: original file path and extracted file path
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: AsposeCellsExamples <originalFilePath> <extractedFilePath>");
                return;
            }

            string originalFilePath = args[0];
            string extractedFilePath = args[1];

            VerifyExtractedFileExtension.Run(originalFilePath, extractedFilePath);
        }
    }
}