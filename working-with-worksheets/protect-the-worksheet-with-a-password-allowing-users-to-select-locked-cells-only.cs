using System;
using Aspose.Cells;

class ProtectWorksheetDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Access the worksheet's protection settings
        Protection protection = sheet.Protection;

        // Allow users to select only locked cells
        protection.AllowSelectingLockedCell = true;
        protection.AllowSelectingUnlockedCell = false; // optional, default is false

        // Set the password for protection
        protection.Password = "mySecretPwd";

        // Protect the worksheet with all protection options
        sheet.Protect(ProtectionType.All);

        // Save the workbook
        workbook.Save("ProtectedSelectLockedOnly.xlsx");
    }
}