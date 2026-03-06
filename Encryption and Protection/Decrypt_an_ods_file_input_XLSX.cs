using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace DecryptOdsExample
{
    class Program
    {
        static void Main()
        {
            // Path to the encrypted ODS file
            string odsPath = "encrypted_file.ods";

            if (!File.Exists(odsPath))
            {
                Console.WriteLine($"File not found: {Path.GetFullPath(odsPath)}");
                return;
            }

            // Password used to protect the ODS file
            string password = "your_password";

            // Detect the file format and whether it is encrypted
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(odsPath);
            Console.WriteLine($"Detected format: {formatInfo.FileFormatType}");
            Console.WriteLine($"Is encrypted: {formatInfo.IsEncrypted}");

            // Create OdsLoadOptions and set the password for decryption
            OdsLoadOptions loadOptions = new OdsLoadOptions
            {
                Password = password
            };

            // Load the ODS workbook using the load options
            Workbook workbook = new Workbook(odsPath, loadOptions);

            // Save the decrypted workbook as XLSX
            string xlsxPath = "decrypted_output.xlsx";
            workbook.Save(xlsxPath, SaveFormat.Xlsx);

            Console.WriteLine($"Decrypted file saved as: {xlsxPath}");
        }
    }
}