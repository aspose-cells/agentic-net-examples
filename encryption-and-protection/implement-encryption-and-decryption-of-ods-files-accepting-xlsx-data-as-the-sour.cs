using System;
using Aspose.Cells;
using Aspose.Cells.Ods;
using Aspose.Cells.Utility;

class OdsEncryptionDemo
{
    static void Main()
    {
        // Input XLSX, encrypted ODS output, and decrypted XLSX output paths
        string sourceXlsx = "source.xlsx";
        string encryptedOds = "encrypted.ods";
        string decryptedXlsx = "decrypted.xlsx";

        // Create a sample XLSX file if it does not exist
        if (!System.IO.File.Exists(sourceXlsx))
        {
            Workbook sampleWb = new Workbook();
            sampleWb.Worksheets[0].Cells["A1"].PutValue("Hello World");
            sampleWb.Worksheets[0].Cells["B1"].PutValue(12345);
            sampleWb.Save(sourceXlsx, SaveFormat.Xlsx);
        }

        // Load the source XLSX workbook
        Workbook wb = new Workbook(sourceXlsx);

        // Set a password – this will encrypt the workbook when saved
        wb.Settings.Password = "myPassword";

        // Configure ODS save options (optional: set generator type)
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        saveOptions.GeneratorType = OdsGeneratorType.LibreOffice;

        // Save the workbook as an encrypted ODS file
        wb.Save(encryptedOds, saveOptions);

        // Detect whether the saved ODS file is encrypted
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(encryptedOds);
        Console.WriteLine("Is the ODS file encrypted? " + formatInfo.IsEncrypted);

        // Load the encrypted ODS file using the password
        OdsLoadOptions loadOptions = new OdsLoadOptions();
        loadOptions.Password = "myPassword";
        Workbook encryptedWb = new Workbook(encryptedOds, loadOptions);

        // Save the decrypted content back to XLSX
        encryptedWb.Save(decryptedXlsx, SaveFormat.Xlsx);

        Console.WriteLine("Decrypted XLSX file saved as: " + decryptedXlsx);
    }
}