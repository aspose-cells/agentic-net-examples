// Title: Encrypt a New Excel Workbook with AES‑256 and a Custom Password using Aspose.Cells for .NET (C#)
// Description: Shows how to create a Workbook with Aspose.Cells, assign a user‑defined password, enable AES‑256 encryption via StrongCryptographicProvider, and save the result as a protected .xlsx file.
// Keywords: Aspose.Cells | AES-256 encryption | C# .NET | Workbook password | StrongCryptographicProvider | Encrypt Excel file | SetEncryptionOptions | Secure Excel export
// Common Searches: Aspose.Cells AES 256 encrypt workbook C# | How to password protect Excel with Aspose.Cells .NET | Set custom password and strong encryption Aspose.Cells | Encrypt newly created workbook Aspose.Cells | C# code for AES‑256 Excel protection
// Developer Intent: Apply AES‑256 encryption and a user‑defined password to a newly created workbook.
// Use Cases: Distribute confidential reports that can only be opened with a known password. | Store financial or personal data in an Excel file that meets corporate security standards. | Automate generation of password‑protected spreadsheets for external partners.
// AI Prompts: Generate C# code that creates an Aspose.Cells workbook, sets a password, applies AES‑256 encryption, and saves the file. | Provide an example with try‑catch blocks for handling errors while encrypting and saving a password‑protected workbook in Aspose.Cells. | Explain how to read the encryption type and key length of an existing encrypted Excel file using Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to create a Workbook with Aspose.Cells, assign a user‑defined password, enable AES‑256 encryption via StrongCryptographicProvider, and save the result as a protected .xlsx file.
class EncryptWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data (optional)
        workbook.Worksheets[0].Cells["A1"].PutValue("Sensitive data");

        // Set the password required to open the workbook
        workbook.Settings.Password = "MyCustomPassword";

        // Apply AES‑256 encryption (StrongCryptographicProvider with 256‑bit key)
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

        // Save the encrypted workbook
        workbook.Save("EncryptedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
