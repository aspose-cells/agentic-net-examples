using System;
using Aspose.Cells;
using Aspose.Cells.Ods;
using Aspose.Cells.Utility;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Path for the encrypted ODS file
            string encryptedOdsPath = "EncryptedWorkbook.ods";

            // ------------------- Create and encrypt ODS -------------------
            // Create a new workbook and add some data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");

            // Set a password to protect the workbook
            workbook.Settings.Password = "StrongPassword123";

            // Optional: define encryption algorithm (ignored for OOXML, but kept for completeness)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save as ODS using OdsSaveOptions
            OdsSaveOptions saveOptions = new OdsSaveOptions();
            workbook.Save(encryptedOdsPath, saveOptions);

            // ------------------- Verify encryption status -------------------
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(encryptedOdsPath);
            Console.WriteLine($"Is the file encrypted? {formatInfo.IsEncrypted}");

            // ------------------- Load (decrypt) the ODS file -------------------
            // Create load options with the password
            OdsLoadOptions loadOptions = new OdsLoadOptions();
            loadOptions.Password = "StrongPassword123";

            // Load the encrypted workbook
            Workbook loadedWorkbook = new Workbook(encryptedOdsPath, loadOptions);

            // Read and display the previously protected cell value
            string cellValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine($"Decrypted cell value: {cellValue}");

            // ------------------- Optional: Convert ODS to XLSX -------------------
            string convertedXlsxPath = "ConvertedWorkbook.xlsx";
            ConversionUtility.Convert(encryptedOdsPath, loadOptions, convertedXlsxPath, new OoxmlSaveOptions());

            Console.WriteLine($"File converted to XLSX: {convertedXlsxPath}");
        }
    }
}