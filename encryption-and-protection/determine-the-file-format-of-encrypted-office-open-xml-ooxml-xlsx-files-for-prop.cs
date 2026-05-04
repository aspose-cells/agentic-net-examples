using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class DetectEncryptedXlsxFormat
    {
        static void Main()
        {
            string encryptedFilePath = "encrypted.xlsx";
            string password = "myPassword";

            FileFormatInfo formatInfo;
            using (FileStream fs = new FileStream(encryptedFilePath, FileMode.Open, FileAccess.Read))
            {
                formatInfo = FileFormatUtil.DetectFileFormat(fs);
            }

            Console.WriteLine($"File Format Type: {formatInfo.FileFormatType}");
            Console.WriteLine($"Is Encrypted: {formatInfo.IsEncrypted}");

            bool passwordValid;
            using (FileStream fs = new FileStream(encryptedFilePath, FileMode.Open, FileAccess.Read))
            {
                passwordValid = FileFormatUtil.VerifyPassword(fs, password);
            }
            Console.WriteLine($"Password valid: {passwordValid}");

            if (passwordValid)
            {
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto)
                {
                    Password = password
                };
                Workbook workbook = new Workbook(encryptedFilePath, loadOptions);
                Console.WriteLine($"Worksheets count: {workbook.Worksheets.Count}");
            }
        }
    }
}