using System;
using Aspose.Cells;

class EncryptWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        workbook.Worksheets[0].Cells["A1"].PutValue("Sensitive Data");

        // Set the password required to open the workbook
        workbook.Settings.Password = "OpenPassword123";

        // Specify encryption options (type is ignored for .xlsx, but shown for completeness)
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Protect workbook structure and windows with a separate password
        workbook.Protect(ProtectionType.All, "StructurePassword456");

        // Save the workbook; for .xlsx the encryption will be SHA‑AES (same as Excel)
        workbook.Save("EncryptedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}