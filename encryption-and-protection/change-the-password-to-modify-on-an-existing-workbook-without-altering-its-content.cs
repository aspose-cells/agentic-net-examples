using System;
using Aspose.Cells;

class ChangeWriteProtectionPassword
{
    static void Main()
    {
        // Path to the existing workbook
        string inputPath = "ExistingWorkbook.xlsx";
        // Path for the updated workbook
        string outputPath = "Workbook_With_NewWritePassword.xlsx";

        // Load the workbook (if the file itself is encrypted, set LoadOptions.Password accordingly)
        Workbook wb = new Workbook(inputPath);

        // Change the write‑protection password (password required to modify the file)
        wb.Settings.WriteProtection.Password = "NewWritePassword123";

        // Save the workbook – content remains unchanged, only the write‑protection password is updated
        wb.Save(outputPath);
    }
}