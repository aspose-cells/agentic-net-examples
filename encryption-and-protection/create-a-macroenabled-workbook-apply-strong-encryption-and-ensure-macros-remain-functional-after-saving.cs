// Title: Create a Macro‑Enabled .xlsm Workbook with AES‑128 Encryption and Password Protection using Aspose.Cells for .NET (C#)
// Description: This C# example demonstrates how to generate a new Workbook, enable macros, assign a strong password, apply AES‑128 encryption via the StrongCryptographicProvider, and save the result as a macro‑enabled .xlsm file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | macro-enabled workbook | xlsm encryption | AES-128 | StrongCryptographicProvider | password protection | EnableMacros | SetEncryptionOptions | secure Excel file
// Common Searches: Aspose.Cells encrypt xlsm file | Enable macros and set password with Aspose.Cells .NET | Save macro-enabled workbook with AES encryption using C# | Create password‑protected .xlsm using Aspose.Cells | Apply StrongCryptographicProvider to macro workbook
// Developer Intent: Produce a .xlsm file that contains VBA macros, is secured with a password, and uses AES‑128 encryption, while ensuring the macros remain executable after saving.
// Use Cases: Distribute a template that includes VBA macros and must be protected against unauthorized access. | Automate generation of encrypted macro‑enabled financial reports that comply with corporate security policies. | Archive user‑generated spreadsheets with macros in a secure repository, requiring a password to open. | Deliver confidential engineering calculations in a macro‑enabled workbook that cannot be tampered with.
// AI Prompts: Show me how to add a VBA module to the workbook before saving it as an encrypted .xlsm with Aspose.Cells. | Explain how to verify that macros still run after applying AES‑128 encryption to a macro‑enabled file. | Provide C# code to switch the encryption strength from 128‑bit to 256‑bit for a macro‑enabled workbook using Aspose.Cells.

using System;
using Aspose.Cells;

// This C# example demonstrates how to generate a new Workbook, enable macros, assign a strong password, apply AES‑128 encryption via the StrongCryptographicProvider, and save the result as a macro‑enabled .xlsm file with Aspose.Cells for .NET.
class MacroEnabledEncryptedWorkbook
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Enable macros in the workbook
            wb.Settings.EnableMacros = true;

            // NOTE: Adding VBA modules requires the VBA component of Aspose.Cells.
            // If the component is unavailable, the workbook will still be saved as a macro‑enabled file.

            // Set a password to protect (encrypt) the workbook
            wb.Settings.Password = "StrongPassword123";

            // Apply strong encryption (AES 128‑bit)
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the workbook as a macro‑enabled file (.xlsm)
            string outputPath = "EncryptedMacroWorkbook.xlsm";
            wb.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
