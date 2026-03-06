using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFileFormatDetection
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = "sample.xlsx";

            // Detect the file format and encryption status
            FileFormatInfo formatInfo;
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                formatInfo = FileFormatUtil.DetectFileFormat(fs);
            }

            Console.WriteLine($"Detected File Format Type: {formatInfo.FileFormatType}");
            Console.WriteLine($"Is Encrypted: {formatInfo.IsEncrypted}");

            // If encrypted, verify the password
            if (formatInfo.IsEncrypted)
            {
                string password = "yourPassword";
                bool passwordValid;
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    passwordValid = FileFormatUtil.VerifyPassword(fs, password);
                }
                Console.WriteLine($"Password valid: {passwordValid}");
            }
        }
    }
}