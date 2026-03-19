using System;
using Aspose.Cells;

class WorkbookPasswordDemo
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sensitive data");

        // Apply workbook protection with a password
        string password = "Secret123";
        wb.Protect(ProtectionType.All, password);

        // Save the password‑protected workbook
        string protectedPath = "ProtectedWorkbook.xlsx";
        wb.Save(protectedPath, SaveFormat.Xlsx);

        // Load the protected workbook using LoadOptions that contain the password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = password;
        Workbook protectedWb = new Workbook(protectedPath, loadOptions);

        // Verify that the workbook is still protected
        Console.WriteLine("Is workbook protected with password: " + protectedWb.IsWorkbookProtectedWithPassword);

        // Remove the protection
        protectedWb.Unprotect(password);

        // Save the unprotected workbook
        string unprotectedPath = "UnprotectedWorkbook.xlsx";
        protectedWb.Save(unprotectedPath, SaveFormat.Xlsx);

        // Load the unprotected workbook to confirm removal of protection
        Workbook finalWb = new Workbook(unprotectedPath);
        Console.WriteLine("Is workbook protected after unprotect: " + finalWb.IsWorkbookProtectedWithPassword);
    }
}