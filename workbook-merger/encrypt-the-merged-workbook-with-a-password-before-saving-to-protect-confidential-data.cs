// Title: Password‑protect and encrypt a merged workbook with Aspose.Cells for .NET (C#)
// Description: Shows how to assign a password, enable 128‑bit strong encryption, and save a merged workbook as an encrypted .xlsx file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | encrypt workbook | password protection | SetEncryptionOptions | EncryptionType.StrongCryptographicProvider | save encrypted Excel | merged workbook security
// Common Searches: Aspose.Cells password protect workbook C# | Encrypt Excel file with Aspose.Cells .NET | Set 128‑bit encryption Aspose.Cells | Secure merged workbook Aspose.Cells | C# code to encrypt Excel after merging
// Developer Intent: Apply a password and strong encryption to a workbook generated from merging multiple Excel files before saving it.
// Use Cases: After consolidating several source workbooks, protect the final file with a password and 128‑bit encryption to meet data‑privacy policies. | Automate report generation that outputs a confidential Excel workbook, ensuring it cannot be opened without the specified password. | Store merged financial statements on a shared drive while guaranteeing that only authorized users can decrypt the file.
// AI Prompts: Generate C# code that merges multiple Excel files with Aspose.Cells, then encrypts the resulting workbook using a password and 128‑bit AES. | Explain how EncryptionType.StrongCryptographicProvider differs from other encryption options in Aspose.Cells for .NET. | Provide step‑by‑step instructions to password‑protect and save an encrypted workbook after a merge operation using Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to assign a password, enable 128‑bit strong encryption, and save a merged workbook as an encrypted .xlsx file using Aspose.Cells for .NET.
class EncryptWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook (this would be the merged workbook in a real scenario)
        Workbook workbook = new Workbook();

        // Add sample data to illustrate the workbook content
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Confidential Data");

        // Set the password that encrypts the workbook when it is saved
        workbook.Settings.Password = "StrongPassword123";

        // (Optional) Define stronger encryption options for Excel 2007/2010+ files
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook to disk
        workbook.Save("EncryptedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
