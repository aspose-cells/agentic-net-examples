// Title: Set an opening password and AES‑128 encryption on a new XLSX workbook with Aspose.Cells for .NET
// Description: Shows how to create a Workbook, add data, assign an opening password, configure AES‑128 encryption via the StrongCryptographicProvider, and save the result as a protected XLSX file.
// Keywords: Aspose.Cells password protection C# | AES 128 encryption Aspose.Cells | secure Excel file .NET | Workbook.SetEncryptionOptions example | protect XLSX with password Aspose | C# encrypt Excel workbook
// Common Searches: how to add a password to an XLSX file using Aspose.Cells | enable AES‑128 encryption for a workbook in C# | Aspose.Cells set opening password .NET | encrypt Excel workbook with strong cryptography Aspose | verify password protection on saved XLSX Aspose.Cells
// Developer Intent: The developer needs to generate an XLSX workbook, protect it with an opening password, and apply strong AES‑128 encryption before writing the file to disk.
// Use Cases: Distribute confidential financial reports that only authorized recipients can open. | Create template files containing proprietary formulas and lock them with robust encryption. | Automate export of sensitive data to meet GDPR or HIPAA compliance requirements.
// AI Prompts: Generate C# code that applies an opening password and AES‑256 encryption to an existing Aspose.Cells workbook. | Explain how to programmatically confirm that a saved XLSX file is encrypted with the specified password using Aspose.Cells. | Provide a snippet to change or remove the password of an already encrypted workbook with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsPasswordEncryptionDemo
{
    // Shows how to create a Workbook, add data, assign an opening password, configure AES‑128 encryption via the StrongCryptographicProvider, and save the result as a protected XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive data");

            // Set the password required to open the workbook
            workbook.Settings.Password = "StrongPassword!123";

            // Enforce strong encryption (AES 128-bit)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook (save rule)
            workbook.Save("EncryptedWorkbook.xlsx", SaveFormat.Xlsx);

            // Optional: Verify that the workbook is encrypted
            Console.WriteLine("Workbook saved with password protection and strong encryption.");
        }
    }
}
