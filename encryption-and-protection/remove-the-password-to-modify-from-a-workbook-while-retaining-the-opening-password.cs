// Title: Remove Excel Write‑Protection Password While Keeping Opening Encryption Password with Aspose.Cells for .NET (C#)
// Description: Loads an encrypted workbook using its opening password, clears the write‑protection (modify) password via Workbook.Settings.WriteProtection.Password, and saves the file so it stays encrypted but no longer prompts for a modify password.
// Keywords: Aspose.Cells remove write protection | C# clear modify password Excel | keep opening password | Excel encryption password | Workbook.Settings.WriteProtection | Aspose.Cells .NET password handling | remove password to modify | preserve file encryption | Excel file security | remove write‑protection programmatically
// Common Searches: Aspose.Cells remove modify password C# | how to keep opening password after clearing write protection | clear write protection password in encrypted Excel workbook | C# Aspose.Cells delete password to modify | remove write‑protection while preserving encryption
// Developer Intent: Strip the modify (write‑protection) password from an encrypted Excel workbook without affecting its opening password.
// Use Cases: Distribute a confidential template that can be opened by recipients but does not require a separate modify password. | Automate batch processing of encrypted workbooks to remove write‑protection before applying further data transformations. | Create a secure Excel file for archiving that remains encrypted yet can be edited by downstream tools without extra prompts.
// AI Prompts: Generate C# code using Aspose.Cells to load an encrypted workbook, clear the write‑protection password, and save it while preserving the opening password. | Explain why setting Workbook.Settings.WriteProtection.Password to null removes the modify password but leaves the encryption password intact. | Write a script that iterates over a folder of Excel files, removes each file's write‑protection password, and saves them with their original opening passwords using Aspose.Cells.

using Aspose.Cells;

// Loads an encrypted workbook using its opening password, clears the write‑protection (modify) password via Workbook.Settings.WriteProtection.Password, and saves the file so it stays encrypted but no longer prompts for a modify password.
class RemoveWriteProtection
{
    static void Main()
    {
        // Load the workbook that is protected with an opening (encryption) password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "OpenPassword"; // opening password
        Workbook workbook = new Workbook("ProtectedWorkbook.xlsx", loadOptions);

        // Remove the "password to modify" (write‑protection password) while keeping the opening password intact
        workbook.Settings.WriteProtection.Password = null; // or string.Empty

        // Save the workbook; it will still require the opening password to be opened
        workbook.Save("WorkbookWithoutModifyPassword.xlsx");
    }
}
