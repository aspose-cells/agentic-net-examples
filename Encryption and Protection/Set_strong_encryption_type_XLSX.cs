using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // (Optional) Add some data to the first worksheet
        workbook.Worksheets[0].Cells["A1"].PutValue("Encrypted content");

        // Set the password required to open the workbook
        workbook.Settings.Password = "StrongPassword123";

        // Apply strong encryption (AES 128‑bit) to the workbook
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the workbook as an XLSX file
        workbook.Save("StrongEncryptedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}