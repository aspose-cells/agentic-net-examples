// Title: Update Password of an Encrypted Excel Workbook using Aspose.Cells for .NET (C#)
// Description: Shows how to open a password‑protected .xlsx by supplying the existing password through LoadOptions, set a stronger credential via Workbook.Settings.Password, and save the workbook so it is re‑encrypted with the new key.
// Keywords: Aspose.Cells password change | C# encrypted Excel workbook | replace Excel file password | load workbook with password | save workbook with new password | Workbook.Settings.Password | Excel re‑encryption .NET
// Common Searches: change Excel file password Aspose.Cells C# | update workbook encryption password .NET | load protected .xlsx using old password Aspose | save encrypted workbook with a new password | upgrade weak Excel password with Aspose.Cells
// Developer Intent: Swap the current workbook password for a stronger one programmatically.
// Use Cases: Upgrade legacy Excel files that use weak passwords to meet modern security standards. | Automate periodic password rotation for confidential workbooks in a maintenance pipeline. | Re‑encrypt a workbook after a policy change while preserving all sheet‑level protections.
// AI Prompts: Write C# code that opens a password‑protected Excel file with Aspose.Cells, changes the password to a stronger value, and saves the file with error handling. | Explain how to confirm that the new password protects the workbook after saving using Aspose.Cells. | Provide a snippet that updates the workbook password without affecting existing sheet protection or custom properties.

using Aspose.Cells;

// Shows how to open a password‑protected .xlsx by supplying the existing password through LoadOptions, set a stronger credential via Workbook.Settings.Password, and save the workbook so it is re‑encrypted with the new key.
class ChangeWorkbookPassword
{
    static void Main()
    {
        // Paths to the source and destination files
        string sourceFile = "protected.xlsx";
        string destinationFile = "protected_newpwd.xlsx";

        // Old (current) password and the new stronger password
        string oldPassword = "oldPassword123";
        string newPassword = "NewStrongPassword!@#2026";

        // Load the encrypted workbook using the old password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = oldPassword;
        Workbook workbook = new Workbook(sourceFile, loadOptions);

        // Update the workbook encryption password to the new value
        workbook.Settings.Password = newPassword;

        // Save the workbook; it will be encrypted with the new password
        workbook.Save(destinationFile);
    }
}
