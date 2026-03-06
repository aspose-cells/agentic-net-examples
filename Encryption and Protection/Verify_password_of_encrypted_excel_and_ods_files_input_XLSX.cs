using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPasswordVerification
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] files = { "encrypted.xlsx", "encrypted.ods", "plain.xlsx" };
            string password = "test";

            foreach (string file in files)
            {
                Console.WriteLine($"--- Verifying \"{file}\" ---");
                VerifyFilePassword(file, password);
                Console.WriteLine();
            }
        }

        static void VerifyFilePassword(string filePath, string password)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
            Console.WriteLine($"IsEncrypted (detected): {formatInfo.IsEncrypted}");

            if (!formatInfo.IsEncrypted)
            {
                Console.WriteLine("File is not encrypted; password verification not required.");
                return;
            }

            bool isValid;
            using (Stream stream = File.OpenRead(filePath))
            {
                isValid = FileFormatUtil.VerifyPassword(stream, password);
            }
            Console.WriteLine($"Password \"{password}\" is valid: {isValid}");

            if (isValid)
            {
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto) { Password = password };
                using (Workbook wb = new Workbook(filePath, loadOptions))
                {
                    Console.WriteLine($"Workbook loaded successfully. First sheet name: {wb.Worksheets[0].Name}");
                }
            }
        }
    }
}