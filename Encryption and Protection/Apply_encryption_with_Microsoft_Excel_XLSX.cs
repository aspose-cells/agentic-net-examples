using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        workbook.Worksheets[0].Cells["A1"].PutValue("Sensitive Information");

        // Set a password – this encrypts the workbook when saved as .xlsx
        workbook.Settings.Password = "StrongPassword123";

        // Set encryption options (for older formats; for .xlsx this is ignored but shown for completeness)
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook
        workbook.Save("EncryptedWorkbook.xlsx", SaveFormat.Xlsx);

        // Verify that the file is encrypted
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat("EncryptedWorkbook.xlsx");
        Console.WriteLine("Is the workbook encrypted? " + formatInfo.IsEncrypted);
    }
}