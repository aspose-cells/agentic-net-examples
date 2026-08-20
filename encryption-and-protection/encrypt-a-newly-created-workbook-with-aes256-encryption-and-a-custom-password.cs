// Title: Encrypt a New Excel Workbook with AES‑256 and Custom Password using Aspose.Cells for .NET
// Description: Creates a workbook, optionally adds data, sets an opening password, applies AES‑256 encryption via StrongCryptographicProvider, and saves the file as an .xlsx. Shows the Aspose.Cells C# API for secure workbook generation.
// Keywords: Aspose.Cells AES-256 encryption C# | workbook password protection .NET | SetEncryptionOptions Aspose.Cells | StrongCryptographicProvider Excel encryption | encrypt Excel file programmatically | C# secure Excel workbook | Aspose.Cells encryption API | AES-256 Excel protection
// Common Searches: How to encrypt an Excel workbook with AES‑256 using Aspose.Cells C# | Set password and encryption type for Aspose.Cells workbook | Aspose.Cells StrongCryptographicProvider example | Encrypt newly created workbook programmatically .NET | Excel file password protection Aspose.Cells
// Developer Intent: Apply AES‑256 encryption and a custom password to a newly created workbook with Aspose.Cells.
// Use Cases: Generate confidential reports that must be protected before distribution. | Create Excel templates programmatically while meeting compliance‑driven encryption standards. | Automate batch production of sensitive workbooks for secure archival storage.
// AI Prompts: Provide C# code to encrypt an Aspose.Cells workbook with AES‑256 and a custom password. | Explain how to use StrongCryptographicProvider with a 256‑bit key to protect an Excel file in Aspose.Cells. | Show steps to verify that a saved workbook is encrypted with AES‑256 using Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, optionally adds data, sets an opening password, applies AES‑256 encryption via StrongCryptographicProvider, and saves the file as an .xlsx. Shows the Aspose.Cells C# API for secure workbook generation.
class EncryptWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data (optional)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sensitive data");

        // Set the password required to open the workbook
        workbook.Settings.Password = "MyStrongPassword";

        // Apply AES‑256 encryption (StrongCryptographicProvider with 256‑bit key)
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

        // Save the encrypted workbook
        workbook.Save("EncryptedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
