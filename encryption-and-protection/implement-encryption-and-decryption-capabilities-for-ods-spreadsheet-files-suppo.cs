using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a new workbook and add sample data
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");
            sheet.Cells["B2"].PutValue(12345);

            // -----------------------------------------------------------------
            // 2. Apply password protection and encryption options
            // -----------------------------------------------------------------
            // Set the password that will protect the workbook
            workbook.Settings.Password = "SecretPwd123";

            // Define encryption algorithm and key length (optional, but recommended)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // -----------------------------------------------------------------
            // 3. Save the workbook as an encrypted ODS file
            // -----------------------------------------------------------------
            OdsSaveOptions odsSaveOptions = new OdsSaveOptions();
            odsSaveOptions.GeneratorType = OdsGeneratorType.LibreOffice; // optional generator setting
            string odsPath = "EncryptedWorkbook.ods";
            workbook.Save(odsPath, odsSaveOptions);

            // -----------------------------------------------------------------
            // 4. Also save the same workbook as an encrypted XLSX file
            // -----------------------------------------------------------------
            string xlsxPath = "EncryptedWorkbook.xlsx";
            workbook.Save(xlsxPath, SaveFormat.Xlsx); // password already applied via Settings.Password

            // -----------------------------------------------------------------
            // 5. Load the encrypted ODS file using the password
            // -----------------------------------------------------------------
            OdsLoadOptions odsLoadOptions = new OdsLoadOptions();
            odsLoadOptions.Password = "SecretPwd123";
            Workbook loadedOds = new Workbook(odsPath, odsLoadOptions);
            Console.WriteLine("ODS Loaded Cell A1: " + loadedOds.Worksheets[0].Cells["A1"].Value);

            // -----------------------------------------------------------------
            // 6. Load the encrypted XLSX file using the password
            // -----------------------------------------------------------------
            LoadOptions xlsxLoadOptions = new LoadOptions();
            xlsxLoadOptions.Password = "SecretPwd123";
            Workbook loadedXlsx = new Workbook(xlsxPath, xlsxLoadOptions);
            Console.WriteLine("XLSX Loaded Cell B2: " + loadedXlsx.Worksheets[0].Cells["B2"].Value);

            // -----------------------------------------------------------------
            // 7. Detect encryption status of the saved files
            // -----------------------------------------------------------------
            var odsInfo = FileFormatUtil.DetectFileFormat(odsPath);
            Console.WriteLine($"ODS IsEncrypted: {odsInfo.IsEncrypted}");

            var xlsxInfo = FileFormatUtil.DetectFileFormat(xlsxPath);
            Console.WriteLine($"XLSX IsEncrypted: {xlsxInfo.IsEncrypted}");

            // Detect with password (required for encrypted OOXML files)
            var xlsxInfoWithPwd = FileFormatUtil.DetectFileFormat(xlsxPath, "SecretPwd123");
            Console.WriteLine($"XLSX IsEncrypted (with password): {xlsxInfoWithPwd.IsEncrypted}");
        }
    }
}