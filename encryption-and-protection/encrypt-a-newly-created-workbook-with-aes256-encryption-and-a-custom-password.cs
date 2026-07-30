// Title: Encrypt a new Excel workbook with AES‑256 and a custom password using Aspose.Cells for .NET (C#)
// Description: Creates a Workbook, writes data to cell A1, assigns a user‑defined password, applies AES‑256 encryption via StrongCryptographicProvider, and saves the file as EncryptedWorkbook.xlsx.
// Keywords: Aspose.Cells | AES-256 encryption | Excel password protection | StrongCryptographicProvider | C# workbook encryption | Save encrypted XLSX
// Common Searches: Aspose.Cells AES-256 encryption C# | set password for Excel file using Aspose | encrypt newly created workbook Aspose.Cells .NET | StrongCryptographicProvider example
// Developer Intent: Apply AES‑256 encryption with a custom password to a newly created workbook.
// Use Cases: Secure financial statements before distribution. | Store confidential HR data at rest to meet compliance standards. | Automate generation of password‑protected reports for external partners.
// AI Prompts: Generate C# code that creates an Aspose.Cells workbook, sets a custom password, applies AES‑256 encryption with StrongCryptographicProvider, and saves it. | Explain step‑by‑step how to configure AES‑256 encryption and a password for a new workbook in Aspose.Cells for .NET. | Show how to verify that an Excel file saved with Aspose.Cells is encrypted with AES‑256 and requires a password to open.

using System;
using Aspose.Cells;

// Creates a Workbook, writes data to cell A1, assigns a user‑defined password, applies AES‑256 encryption via StrongCryptographicProvider, and saves the file as EncryptedWorkbook.xlsx.
class EncryptWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Add some data to the first worksheet
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sensitive data");

        // Set the password required to open the workbook
        wb.Settings.Password = "MyStrongPassword";

        // Apply AES‑256 encryption (StrongCryptographicProvider with 256‑bit key)
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

        // Save the encrypted workbook
        wb.Save("EncryptedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
