using Aspose.Cells;
using Aspose.Cells.Ods;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add some data to the first worksheet
        workbook.Worksheets[0].Cells["A1"].PutValue("Sensitive information");

        // Set a password – this will encrypt the file when saved
        workbook.Settings.Password = "StrongPassword123";

        // Optional: specify encryption algorithm (ignored for ODS but harmless)
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Create ODS save options (default settings are sufficient)
        OdsSaveOptions odsOptions = new OdsSaveOptions();

        // Save the workbook as an encrypted ODS file
        workbook.Save("EncryptedDocument.ods", odsOptions);
    }
}