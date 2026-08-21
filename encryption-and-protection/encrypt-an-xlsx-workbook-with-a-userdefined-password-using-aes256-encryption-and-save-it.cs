// Title: Encrypt an XLSX workbook with a password using AES‑256 in Aspose.Cells for .NET
// Description: Creates a new Workbook, writes text to cell A1, assigns a user‑defined password, applies AES‑256 encryption via StrongCryptographicProvider, and saves the result as EncryptedWorkbook.xlsx.
// Keywords: Aspose.Cells | AES-256 | C# encryption | Excel password protection | StrongCryptographicProvider | SetEncryptionOptions | Workbook.Settings.Password | Encrypt XLSX | .NET
// Common Searches: Aspose.Cells AES-256 encryption example C# | How to password‑protect an XLSX file with Aspose.Cells | Set encryption options for Excel workbook in .NET | Encrypt Excel workbook using StrongCryptographicProvider | C# code to save encrypted XLSX with custom password
// Developer Intent: Apply AES‑256 password protection to an XLSX file using Aspose.Cells.
// Use Cases: Secure financial reports before emailing to clients. | Create compliance‑ready spreadsheets that require a password to open. | Batch‑encrypt archived workbooks with strong encryption for data protection.
// AI Prompts: Show C# code that encrypts an existing workbook with a password using AES‑256 in Aspose.Cells. | Give an example of setting a custom password and enabling StrongCryptographicProvider with a 256‑bit key for an XLSX file. | Explain how to change the encryption algorithm from AES‑256 to AES‑128 with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Creates a new Workbook, writes text to cell A1, assigns a user‑defined password, applies AES‑256 encryption via StrongCryptographicProvider, and saves the result as EncryptedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("This workbook is encrypted with AES‑256.");

        // Set the password that will be required to open the workbook
        workbook.Settings.Password = "MySecretPassword";

        // Configure encryption to use StrongCryptographicProvider (AES) with a 256‑bit key
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

        // Save the encrypted workbook to an XLSX file
        workbook.Save("EncryptedWorkbook.xlsx");
    }
}
