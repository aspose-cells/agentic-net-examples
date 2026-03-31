using System;
using Aspose.Cells;

class PasswordProtectionDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Add some data to the first worksheet
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sensitive Data");

        // ---- Restrict opening ----
        // Set a password that encrypts the file; the workbook cannot be opened without it
        wb.Settings.Password = "openPwd";

        // ---- Restrict structure (adding/removing/renaming sheets) ----
        // Protect the workbook's structure with a password
        wb.Protect(ProtectionType.Structure, "structurePwd");

        // ---- Restrict editing (write protection) ----
        // Set a write‑protection password and recommend read‑only mode
        wb.Settings.WriteProtection.Password = "writePwd";
        wb.Settings.WriteProtection.RecommendReadOnly = true;

        // Save the protected workbook
        string filePath = "ProtectedWorkbook.xlsx";
        wb.Save(filePath, SaveFormat.Xlsx);

        // Load the workbook using the opening password to verify it works
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "openPwd";
        Workbook loadedWb = new Workbook(filePath, loadOptions);

        // Output a cell value to confirm successful load
        Console.WriteLine("Cell A1 value: " + loadedWb.Worksheets[0].Cells["A1"].StringValue);
    }
}