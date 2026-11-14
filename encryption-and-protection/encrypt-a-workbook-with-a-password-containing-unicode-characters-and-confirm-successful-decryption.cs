// Title: Encrypt an Excel workbook with a Unicode (non‑ASCII) password including emoji and verify decryption using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set workbook password with Unicode characters and emoji | load password‑protected xlsx that uses non‑ASCII password in .NET | how to decrypt an Excel file encrypted with a Unicode password using Aspose.Cells | C# example for protecting Excel workbook with accented characters in password | verify workbook decryption after applying Unicode password in Aspose.Cells
// Tags: set workbook password Unicode Aspose.Cells | load encrypted xlsx with Unicode password C# | encrypt Excel file non‑ASCII password Aspose.Cells | verify decryption of password‑protected workbook .NET | Unicode characters in Excel password Aspose.Cells

using System;
using Aspose.Cells;

// Demonstrates how to protect an Excel workbook with a Unicode password that includes accented characters and an emoji, save the encrypted file, then open it using LoadOptions with the same password to confirm successful decryption.
class Program
{
    static void Main()
    {
        // Unicode password (contains accented characters and an emoji)
        string unicodePassword = "Päßwörd🔐";

        // -------------------------
        // Create a new workbook and add sample data
        // -------------------------
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Secret Data");

        // Apply password protection using the Unicode password
        wb.Settings.Password = unicodePassword;

        // Save the encrypted workbook
        string encryptedFile = "encrypted.xlsx";
        wb.Save(encryptedFile);

        // -------------------------
        // Load the encrypted workbook using the same password
        // -------------------------
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            Password = unicodePassword
        };
        Workbook decryptedWb = new Workbook(encryptedFile, loadOptions);

        // Verify decryption by reading the cell value
        string cellValue = decryptedWb.Worksheets[0].Cells["A1"].StringValue;
        Console.WriteLine("Decrypted cell A1 value: " + cellValue);
    }
}
