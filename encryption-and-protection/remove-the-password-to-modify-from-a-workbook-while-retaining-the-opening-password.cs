using System;
using Aspose.Cells;

class RemoveWriteProtectionPassword
{
    static void Main()
    {
        // File path for the workbook
        string filePath = "protected_workbook.xlsx";

        // Password required to open the workbook (encryption password)
        string openingPassword = "open123";

        // Password required to modify the workbook (write‑protection password)
        string modifyPassword = "modify456";

        // ---------- Create a workbook with both passwords ----------
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Sample data");

        // Set the opening (encryption) password
        wb.Settings.Password = openingPassword;

        // Set the write‑protection password (password to modify)
        wb.Settings.WriteProtection.Password = modifyPassword;

        // Save the workbook; it now has both passwords applied
        wb.Save(filePath);
        wb.Dispose();

        // ---------- Load the workbook using the opening password ----------
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = openingPassword; // required to open the file
        Workbook loadedWb = new Workbook(filePath, loadOptions);

        // ---------- Remove the write‑protection password ----------
        // Setting it to null (or string.Empty) clears the "password to modify"
        loadedWb.Settings.WriteProtection.Password = null;

        // Save the workbook again; the opening password remains unchanged
        loadedWb.Save(filePath);
        loadedWb.Dispose();
    }
}