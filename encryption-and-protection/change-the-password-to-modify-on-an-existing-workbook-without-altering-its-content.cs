using System;
using Aspose.Cells;

class ChangeWriteProtectionPassword
{
    static void Main()
    {
        // Path to the existing workbook
        string inputPath = "input.xlsx";

        // Existing write‑protection password (if the file is write‑protected)
        string oldPassword = "oldPass";

        // New password that will replace the old one
        string newPassword = "newPass";

        // Load the workbook with the current password (required for write‑protected files)
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = oldPassword;               // existing password
        Workbook wb = new Workbook(inputPath, loadOptions); // load without changing content

        // Change the write‑protection password (password to modify)
        wb.Settings.WriteProtection.Password = newPassword;

        // Save the workbook – content remains unchanged, only the password is updated
        string outputPath = "output.xlsx";
        wb.Save(outputPath);

        Console.WriteLine($"Write‑protection password changed and saved to '{outputPath}'.");
    }
}