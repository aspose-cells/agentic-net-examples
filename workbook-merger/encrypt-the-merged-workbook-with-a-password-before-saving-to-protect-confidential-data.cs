// Title: Encrypt a merged Excel workbook with a password and 128‑bit strong encryption using Aspose.Cells for .NET
// AI Prompts: Write C# code that builds a merged workbook, assigns an opening password, applies 128‑bit strong encryption, and saves it as an XLSX file with Aspose.Cells. | Refactor the example so the password is passed as a method argument and the encrypted workbook is written to a MemoryStream instead of a physical file.
// Common Searches: Aspose.Cells C# set opening password for merged workbook | apply 128-bit encryption to Excel file using Aspose.Cells .NET | save encrypted workbook to memory stream Aspose.Cells example | protect confidential data in merged Excel workbook with password Aspose.Cells | how to use SetEncryptionOptions with StrongCryptographicProvider in C#
// Tags: Aspose.Cells workbook password protection C# | SetEncryptionOptions strong cryptographic provider Aspose.Cells | encrypt merged Excel workbook Aspose.Cells | save encrypted workbook to memory stream Aspose.Cells | 128-bit encryption Aspose.Cells XLSX

using System;
using Aspose.Cells;

// // Demonstrates creating a merged workbook, inserting confidential data, setting an opening password, applying 128‑bit strong encryption via SetEncryptionOptions, and saving the encrypted file as XLSX using Aspose.Cells for .NET.
class EncryptWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook (this would be the merged workbook)
        Workbook wb = new Workbook();

        // Example: add some confidential data to the first worksheet
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Confidential Data");

        // Set the password that will be required to open the workbook
        wb.Settings.Password = "StrongPassword123";

        // Apply strong encryption (128‑bit key) – optional but recommended
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook to disk
        wb.Save("EncryptedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
