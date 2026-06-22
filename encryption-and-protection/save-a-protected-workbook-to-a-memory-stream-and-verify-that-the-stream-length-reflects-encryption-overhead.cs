using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a workbook and add some sample data
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Sample");
            ws.Cells["B1"].PutValue(123);

            // Save the unprotected workbook to a memory stream
            MemoryStream unprotectedStream = wb.SaveToStream();

            // Protect the workbook with a password
            wb.Settings.Password = "SecretPwd";

            // Optionally set stronger encryption options (ignored for .xlsx but kept for completeness)
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the protected workbook to another memory stream
            MemoryStream protectedStream = wb.SaveToStream();

            // Verify that the protected stream length is greater due to encryption overhead
            Console.WriteLine($"Unprotected stream length: {unprotectedStream.Length} bytes");
            Console.WriteLine($"Protected stream length:   {protectedStream.Length} bytes");
            Console.WriteLine($"Encryption overhead detected: {protectedStream.Length > unprotectedStream.Length}");

            // Additional verification using FileFormatInfo
            // Reset stream position before detection
            protectedStream.Position = 0;
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(protectedStream);
            Console.WriteLine($"FileFormatInfo reports encrypted: {formatInfo.IsEncrypted}");

            // Clean up streams
            unprotectedStream.Dispose();
            protectedStream.Dispose();
        }
    }
}