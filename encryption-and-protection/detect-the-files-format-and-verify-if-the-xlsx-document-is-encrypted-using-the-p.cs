using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath;
            if (args.Length > 0)
                filePath = args[0];
            else
            {
                Console.Write("Enter Excel file path: ");
                filePath = Console.ReadLine() ?? string.Empty;
            }

            DetectAndVerifyEncryption.Run(filePath);
        }
    }

    public static class DetectAndVerifyEncryption
    {
        public static void Run(string filePath)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            using (FileStream fs = File.OpenRead(filePath))
            {
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(fs);
                Console.WriteLine($"Detected File Format Type: {formatInfo.FileFormatType}");
                Console.WriteLine($"Is Encrypted: {formatInfo.IsEncrypted}");

                if (formatInfo.IsEncrypted)
                {
                    Console.Write("Enter password to verify: ");
                    string testPassword = Console.ReadLine() ?? string.Empty;

                    // Reset stream position before verification
                    fs.Position = 0;
                    bool passwordValid = FileFormatUtil.VerifyPassword(fs, testPassword);
                    Console.WriteLine($"Password \"{testPassword}\" valid: {passwordValid}");
                }
            }
        }
    }
}