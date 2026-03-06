using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CheckEncryptionStatus
    {
        public static void Run(string filePath)
        {
            // Detect the file format and retrieve information about the file
            FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);

            // Use the IsEncrypted property to determine if the workbook requires a password
            bool isEncrypted = fileInfo.IsEncrypted;

            Console.WriteLine($"File \"{filePath}\" is encrypted: {isEncrypted}");
        }

        // Example entry point
        public static void Main()
        {
            // Replace with the path to the XLSX file you want to check
            string inputFile = "input.xlsx";

            Run(inputFile);
        }
    }
}