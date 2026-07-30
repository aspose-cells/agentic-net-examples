// Title: Save a protected and encrypted XLSX workbook with Aspose.Cells in C#
// Description: Creates a new Workbook, applies full structure protection, sets an opening password for encryption, and saves the file as XLSX using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# workbook protection | Excel structure password | Excel file encryption | Save encrypted XLSX | Protect workbook with Aspose.Cells | Aspose.Cells SaveFormat.Xlsx
// Common Searches: Aspose.Cells protect workbook structure C# | Set opening password Aspose.Cells .NET | Save encrypted XLSX with Aspose.Cells | How to apply workbook protection and encryption using Aspose.Cells | C# code to protect and encrypt Excel file with Aspose.Cells
// Developer Intent: Developer needs to generate an Excel file, lock its structure, encrypt it with an opening password, and write it to disk as an XLSX file using Aspose.Cells for .NET.
// Use Cases: Distribute read‑only templates where layout cannot be altered without a password. | Send confidential financial statements that require a password to open. | Create data‑entry workbooks where only specific sheets are editable while the overall workbook remains locked. | Automate generation of password‑protected reports for regulatory compliance.
// AI Prompts: Generate C# code that protects selected worksheets and sets a different opening password with Aspose.Cells. | Show how to apply separate passwords for structure protection and file encryption in Aspose.Cells for .NET. | Explain how to programmatically verify workbook protection and encryption after saving the XLSX file. | Provide a step‑by‑step guide to change the opening password of an existing protected workbook using Aspose.Cells.

using Aspose.Cells;

// Creates a new Workbook, applies full structure protection, sets an opening password for encryption, and saves the file as XLSX using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Protect the workbook's structure with a password
        workbook.Protect(ProtectionType.All, "structurePassword");

        // Set the password required to open the workbook (encryption)
        workbook.Settings.Password = "openPassword";

        // Save the protected and encrypted workbook in XLSX format
        workbook.Save("ProtectedEncryptedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
