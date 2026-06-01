using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and add some data
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encrypted ODS test");

            // 2. Set the password that will encrypt the workbook
            string password = "SecretPwd123";
            wb.Settings.Password = password;

            // 3. Save the workbook as ODS (OpenDocument Spreadsheet)
            string filePath = "EncryptedWorkbook.ods";
            wb.Save(filePath, SaveFormat.ODS);

            // 4. Verify that the file is encrypted using FileFormatInfo
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
            Console.WriteLine($"Is the ODS file encrypted? {formatInfo.IsEncrypted}");

            // 5. Load the encrypted workbook with the correct password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = password;
            Workbook loadedWb = new Workbook(filePath, loadOptions);

            // 6. Confirm that the data can be read after providing the password
            string cellValue = loadedWb.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine($"Cell A1 value after decryption: {cellValue}");
        }
    }
}