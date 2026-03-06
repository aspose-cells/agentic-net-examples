using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DetectFormatAndEncryption
    {
        public static void Run()
        {
            // Path to the XLSX file to be examined
            string filePath = "input.xlsx";

            // Detect the file format and retrieve information about the file
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);

            // Output detected file format
            Console.WriteLine($"Detected File Format Type: {formatInfo.FileFormatType}");

            // Check and output whether the file is encrypted
            Console.WriteLine($"Is Encrypted: {formatInfo.IsEncrypted}");

            // Optional: If the file is encrypted, you can verify a password
            if (formatInfo.IsEncrypted)
            {
                // Example password to test
                string password = "testPassword";

                // Verify the password against the encrypted file
                using (Stream stream = File.OpenRead(filePath))
                {
                    bool isPasswordValid = FileFormatUtil.VerifyPassword(stream, password);
                    Console.WriteLine($"Is password \"{password}\" valid: {isPasswordValid}");
                }
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DetectFormatAndEncryption.Run();
        }
    }
}